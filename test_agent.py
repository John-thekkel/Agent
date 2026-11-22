import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import agent
from agent import RUBRIC_WEIGHTS, score_rubric
from docx import Document


class RubricTests(unittest.TestCase):
    def test_quality_gate_requires_grounded_claims_even_with_high_total(self):
        scores = {criterion: 5 for criterion in RUBRIC_WEIGHTS}
        scores["factual_grounding"] = 3

        result = score_rubric({"scores": scores})

        self.assertEqual(result["overall_score"], 88)
        self.assertFalse(result["passed"])

    def test_passing_threshold_and_rubric_validation(self):
        passing = {criterion: 4 for criterion in RUBRIC_WEIGHTS}
        result = score_rubric({"scores": passing, "issues": [], "revision_instructions": ""})
        self.assertEqual(result["overall_score"], 80)
        self.assertTrue(result["passed"])

        passing["factual_grounding"] = True
        with self.assertRaises(ValueError):
            score_rubric({"scores": passing})

    def test_pipeline_revises_once_when_critic_fails(self):
        match = {"match_score": 90, "skip": False}
        first_draft = {"tailored_summary": "First draft"}
        revised_draft = {"tailored_summary": "Revised draft"}
        failed_review = {"passed": False, "overall_score": 70}
        passed_review = {"passed": True, "overall_score": 90}
        with patch.object(agent, "match_job", return_value=match), \
                patch.object(agent, "write_application",
                             side_effect=[first_draft, revised_draft]) as writer, \
                patch.object(agent, "critique_application",
                             side_effect=[failed_review, passed_review]):
            analysis, evaluation, assessment = agent.run_pipeline({"id": "test-job"})

        self.assertEqual(writer.call_count, 2)
        self.assertEqual(analysis["tailored_summary"], "Revised draft")
        self.assertIs(evaluation, passed_review)
        self.assertIs(assessment, match)


class JobSourceTests(unittest.TestCase):
    class Response:
        def __init__(self, data):
            self.data = data

        def raise_for_status(self):
            pass

        def json(self):
            return self.data

    def test_jobicy_and_himalayas_normalize_to_pipeline_shape(self):
        jobicy = agent.normalize_jobicy_job({
            "id": 42,
            "jobTitle": "Python Developer",
            "companyName": "Example",
            "jobGeo": "Remote",
            "salaryMin": 1000,
            "salaryMax": 2000,
            "salaryCurrency": "USD",
            "url": "https://jobicy.com/jobs/42",
            "jobDescription": "&lt;p&gt;Build Python APIs&lt;/p&gt;",
        })
        self.assertEqual(jobicy["id"], "jobicy-42")
        self.assertEqual(jobicy["description"], "Build Python APIs")
        self.assertEqual(jobicy["source"], "Jobicy")

        himalayas = agent.normalize_himalayas_job({
            "guid": "https://himalayas.app/jobs/123",
            "title": "Java Developer",
            "companyName": "Example",
            "locationRestrictions": ["India"],
            "applicationLink": "https://himalayas.app/apply/123",
            "description": "<p>Build Java services</p>",
        })
        self.assertEqual(himalayas["location"], "India")
        self.assertEqual(himalayas["url"], "https://himalayas.app/apply/123")
        self.assertEqual(himalayas["description"], "Build Java services")

    def test_we_work_remotely_rss_is_parsed_and_html_is_cleaned(self):
        feed = """<?xml version="1.0"?>
        <rss><channel><item>
          <title>Example Co: Python Backend Developer</title>
          <link>https://weworkremotely.com/jobs/123</link>
          <region>Anywhere in the World</region>
          <country>India</country>
          <description>&amp;lt;p&amp;gt;Build Python APIs&amp;lt;/p&amp;gt;</description>
        </item></channel></rss>"""

        [job] = agent.parse_weworkremotely_feed(feed)

        self.assertEqual(job["company"], "Example Co")
        self.assertEqual(job["title"], "Python Backend Developer")
        self.assertEqual(job["location"], "Anywhere in the World, India")
        self.assertEqual(job["description"], "Build Python APIs")
        self.assertEqual(job["source"], "We Work Remotely")

    def test_search_relevance_includes_multiple_languages(self):
        for title in ("Python Developer", "Java Backend Engineer", "C++ Developer",
                      ".NET Developer"):
            self.assertTrue(agent.relevant({"title": title}), title)
        self.assertFalse(agent.relevant({"title": "Marketing Manager"}))

    def test_generated_resume_includes_configured_projects(self):
        analysis = {
            "tailored_summary": "Backend developer.",
            "skills_order": ["Python", "C#"],
            "experience_bullets": ["Built REST APIs."],
        }
        profile = {
            "name": "Sample Candidate",
            "location": "Sample Location",
            "phone": "Sample Phone",
            "email": "sample@example.com",
            "linkedin": "https://linkedin.com/in/sample",
            "title": "Sample Developer",
            "company": "Sample Company",
            "dates": "Sample Dates",
            "education": "Sample Education",
            "certifications": [],
            "projects": [{
                "name": "Sample API Project",
                "context": "Demo project",
                "description": "Built a sample API.",
                "technologies": ["Python"],
            }],
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "resume.docx"
            with patch.object(agent, "PROFILE", profile):
                agent.build_resume(path, analysis)
            paragraphs = [paragraph.text for paragraph in Document(path).paragraphs]

        text = "\n".join(paragraphs)
        self.assertIn("Sample API Project", text)
        self.assertIn("Built a sample API.", text)

    def test_jobicy_fetch_uses_search_terms_and_deduplicates(self):
        payload = {"jobs": [{
            "id": 42,
            "jobTitle": "Python Developer",
            "companyName": "Example",
            "url": "https://jobicy.com/jobs/42",
        }]}
        with patch.dict(agent.PREFS, {"search_terms": ["python", "java"]}), \
                patch.object(agent.requests, "get",
                             return_value=self.Response(payload)) as request, \
                patch.object(agent.time, "sleep") as sleep:
            jobs = agent.fetch_jobicy()

        self.assertEqual(len(jobs), 1)
        self.assertEqual(request.call_count, 2)
        self.assertEqual(request.call_args_list[0].kwargs["params"]["tag"], "python")
        self.assertEqual(request.call_args_list[1].kwargs["params"]["tag"], "java")
        sleep.assert_called_once_with(1)

    def test_himalayas_fetch_follows_cursor(self):
        first_page = self.Response({
            "jobs": [{
                "guid": "https://himalayas.app/jobs/123",
                "title": "Python Developer",
                "companyName": "Example",
            }],
            "nextCursor": "next-page",
        })
        final_page = self.Response({"jobs": [], "nextCursor": None})
        with patch.object(agent.requests, "get",
                          side_effect=[first_page, final_page]) as request, \
                patch.object(agent.time, "sleep") as sleep:
            jobs = agent.fetch_himalayas()

        self.assertEqual(len(jobs), 1)
        self.assertEqual(request.call_args_list[1].kwargs["params"]["cursor"], "next-page")
        sleep.assert_called_once_with(1)


if __name__ == "__main__":
    unittest.main()

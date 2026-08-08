import unittest

import app


class AppSmokeTests(unittest.TestCase):
    def setUp(self):
        self.client = app.app.test_client()
        self.headers = {"X-Toolbox-Token": app.SESSION_TOKEN}

    def test_home_renders_ai_assistant(self):
        response = self.client.get("/")
        self.assertEqual(200, response.status_code)
        self.assertIn(b'id="p-ai"', response.data)

    def test_write_api_requires_session_token(self):
        response = self.client.post("/api/ai/audit", json={"sql": "SELECT 1"})
        self.assertEqual(403, response.status_code)

    def test_static_audit_route(self):
        response = self.client.post("/api/ai/audit",
                                    json={"sql": "SELECT id FROM demo LIMIT 1"},
                                    headers=self.headers)
        self.assertEqual(200, response.status_code)
        self.assertTrue(response.json["audit"]["safe"])

    def test_ai_logic_route_exists(self):
        routes = {rule.rule for rule in app.app.url_map.iter_rules()}
        self.assertIn("/api/sql/ai-logic", routes)

    def test_ai_logic_combines_static_evidence_and_model_result(self):
        class FakeAI:
            def compare_sql(self, sql_a, sql_b, evidence, **kwargs):
                self.evidence = evidence
                return {
                    "conclusion": "JOIN semantics changed",
                    "confidence": "high",
                    "evidence_basis": ["LEFT became INNER"],
                    "differences": [], "hypotheses": [],
                    "verification_steps": [], "blind_spots": [],
                }

        original = app.AI_RUNTIME
        fake = FakeAI()
        app.AI_RUNTIME = fake
        try:
            response = self.client.post("/api/sql/ai-logic", json={
                "sql_a": "SELECT a.id FROM orders a LEFT JOIN items b ON a.id=b.id",
                "sql_b": "SELECT a.id FROM orders a INNER JOIN items b ON a.id=b.id",
                "context": "retain all orders",
            }, headers=self.headers)
        finally:
            app.AI_RUNTIME = original
        self.assertEqual(200, response.status_code)
        self.assertTrue(response.json["ok"])
        self.assertEqual("high", response.json["confidence"])
        self.assertIn("joins", fake.evidence)
        self.assertIn("static", response.json)


if __name__ == "__main__":
    unittest.main()

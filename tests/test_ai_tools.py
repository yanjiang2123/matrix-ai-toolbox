import json
import unittest

from ai_tools import AIConfigError, RuntimeAI, audit_readonly_sql


class ReadonlyAuditTests(unittest.TestCase):
    def test_safe_query(self):
        result = audit_readonly_sql("SELECT id, amount FROM orders LIMIT 20")
        self.assertTrue(result["safe"])
        self.assertEqual("safe", result["level"])

    def test_keywords_inside_string_are_not_blocked(self):
        result = audit_readonly_sql("SELECT 'delete from x' AS note LIMIT 1")
        self.assertTrue(result["safe"])

    def test_multiple_statements_are_blocked(self):
        result = audit_readonly_sql("SELECT 1; DROP TABLE orders")
        self.assertFalse(result["safe"])

    def test_write_inside_cte_is_blocked(self):
        result = audit_readonly_sql("WITH x AS (DELETE FROM orders) SELECT * FROM x")
        self.assertFalse(result["safe"])

    def test_mysql_executable_comment_is_blocked(self):
        result = audit_readonly_sql(
            "SELECT 1 /*!50000 INTO OUTFILE '/tmp/export.txt' */"
        )
        self.assertFalse(result["safe"])

    def test_mariadb_executable_comment_is_blocked(self):
        result = audit_readonly_sql(
            "SELECT 1 /*M!100100 INTO OUTFILE '/tmp/export.txt' */"
        )
        self.assertFalse(result["safe"])

    def test_mysql_dash_comment_requires_following_whitespace(self):
        for sql in (
            "SELECT 1--1; DROP TABLE orders",
            "SELECT 1--x; DELETE FROM orders",
        ):
            with self.subTest(sql=sql):
                self.assertFalse(audit_readonly_sql(sql)["safe"])

    def test_hash_operator_cannot_hide_a_second_statement(self):
        self.assertFalse(audit_readonly_sql(
            "SELECT 1#1; DROP TABLE orders"
        )["safe"])

    def test_valid_dash_comment_is_still_accepted(self):
        result = audit_readonly_sql("SELECT 1 -- explanation\nLIMIT 1")
        self.assertTrue(result["safe"])

    def test_executable_comment_identifier_is_not_a_false_positive(self):
        result = audit_readonly_sql(
            "SELECT executable_comment FROM audit_log LIMIT 1"
        )
        self.assertTrue(result["safe"])

    def test_select_star_is_warning_not_block(self):
        result = audit_readonly_sql("SELECT * FROM orders LIMIT 20")
        self.assertTrue(result["safe"])
        self.assertEqual("warning", result["level"])


class RuntimeConfigTests(unittest.TestCase):
    def test_local_http_is_allowed(self):
        state = RuntimeAI().configure({
            "base_url": "http://127.0.0.1:11434/v1",
            "model": "local-model",
        })
        self.assertTrue(state["configured"])
        self.assertFalse(state["has_api_key"])

    def test_external_http_requires_confirmation(self):
        with self.assertRaises(AIConfigError):
            RuntimeAI().configure({
                "base_url": "http://model.example/v1",
                "model": "example-model",
            })

    def test_secret_is_not_returned(self):
        state = RuntimeAI().configure({
            "base_url": "https://model.example/v1",
            "model": "example-model",
            "api_key": "do-not-return-this",
        })
        self.assertTrue(state["has_api_key"])
        self.assertNotIn("api_key", state)


class SqlReasoningTests(unittest.TestCase):
    def test_reasoning_keeps_only_readonly_verification_sql(self):
        class FakeAI(RuntimeAI):
            def chat(self, system, user, max_tokens=1400):
                return json.dumps({
                    "conclusion": "JOIN 变化可能减少结果行",
                    "confidence": "high",
                    "evidence_basis": ["LEFT JOIN 改为 INNER JOIN"],
                    "differences": [{
                        "category": "JOIN 保留规则", "severity": "high",
                        "evidence": "A LEFT JOIN；B INNER JOIN",
                        "impact": "无明细的主表行可能消失",
                        "reasoning": "INNER JOIN 要求右表命中",
                    }],
                    "hypotheses": [],
                    "verification_steps": [
                        {"title": "统计未匹配行", "purpose": "确认影响范围",
                         "sql": "SELECT COUNT(*) FROM orders LIMIT 1"},
                        {"title": "危险建议", "purpose": "不应返回",
                         "sql": "DELETE FROM orders"},
                    ],
                    "blind_spots": ["未读取真实数据"],
                }, ensure_ascii=False)

        result = FakeAI().compare_sql(
            "SELECT a.id FROM orders a LEFT JOIN items b ON a.id=b.id",
            "SELECT a.id FROM orders a INNER JOIN items b ON a.id=b.id",
            {"joins": {"a": ["LEFT"], "b": ["INNER"]}},
        )
        self.assertEqual("high", result["confidence"])
        self.assertTrue(result["verification_steps"][0]["audit"]["safe"])
        self.assertEqual("", result["verification_steps"][1]["sql"])
        self.assertFalse(result["verification_steps"][1]["audit"]["safe"])


if __name__ == "__main__":
    unittest.main()

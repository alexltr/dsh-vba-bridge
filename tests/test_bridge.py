import pathlib
import sys
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "bridge"))

import vba_bridge as bridge  # noqa: E402


class ProcedureParserTests(unittest.TestCase):
    def test_parses_sub_function_and_property(self):
        code = "\r\n".join(
            [
                "Option Explicit",
                "' a comment",
                "Private Sub Hello()",
                "    Debug.Print \"hi\"",
                "End Sub",
                "Public Function Add(a As Long, b As Long) As Long",
                "    Add = a + b",
                "End Function",
                "Public Property Get Name() As String",
                "    Name = \"x\"",
                "End Property",
            ]
        )
        procs = bridge.parse_procedures(code)
        self.assertEqual([p["name"] for p in procs], ["Hello", "Add", "Name"])
        self.assertEqual([p["kind"] for p in procs], ["Sub", "Function", "Property Get"])
        self.assertEqual(procs[0]["scope"], "Private")
        self.assertEqual(procs[1]["args"], "a As Long, b As Long")
        self.assertEqual(procs[0]["start"], 3)
        self.assertEqual(procs[0]["end"], 5)

    def test_joins_continued_signature(self):
        code = "\r\n".join(
            [
                "Private Sub Handler( _",
                "    ByVal a As Long, _",
                "    ByVal b As String)",
                "End Sub",
            ]
        )
        procs = bridge.parse_procedures(code)
        self.assertEqual(len(procs), 1)
        self.assertIn("ByVal a As Long", procs[0]["args"])
        self.assertIn("ByVal b As String", procs[0]["args"])

    def test_ignores_declare_and_attributes(self):
        code = "\r\n".join(
            [
                'Attribute VB_Name = "M"',
                "Private Declare PtrSafe Function GetTickCount Lib \"kernel32\" () As Long",
                "Sub RealOne()",
                "End Sub",
            ]
        )
        self.assertEqual([p["name"] for p in bridge.parse_procedures(code)], ["RealOne"])


class PackagingTests(unittest.TestCase):
    def test_safe_filename(self):
        self.assertEqual(bridge._safe_filename('a/b:c*?"<>|'), "a_b_c______")
        self.assertEqual(bridge._safe_filename("模块1"), "模块1")

    def test_norm_code(self):
        self.assertEqual(bridge.norm_code("a\r\nb\rc"), "a\nb\nc")


if __name__ == "__main__":
    unittest.main()

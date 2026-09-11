import io
import unittest

import pandas as pd

from app import normalize_course, normalize_major, transform_feedback


class FeedbackMergeTests(unittest.TestCase):
    def make_workbook(self, rows):
        columns = {
            "周数": [],
            "您的信息-姓名": [],
            "3、专业": [],
            "班级": [],
            "教师所在院": [],
            "教师信息-姓名": [],
            "教授课程": [],
            "反馈内容": [],
        }
        frame = pd.DataFrame(rows, columns=columns)
        output = io.BytesIO()
        frame.to_excel(output, index=False, engine="openpyxl")
        return output.getvalue()

    def test_major_alias_is_canonicalized(self):
        self.assertEqual(normalize_major("数学应用数学"), "数学与应用数学")
        self.assertEqual(normalize_major(" 数学与应用数学专业 "), "数学与应用数学")
        self.assertEqual(normalize_major("大数据"), "数据科学与大数据技术")
        self.assertEqual(normalize_major("应用统计"), "应用统计学")
        self.assertEqual(normalize_major("金融"), "金融学")

    def test_course_name_completion_is_conservative(self):
        self.assertEqual(normalize_course("统计软件"), "统计软件课程设计")
        self.assertEqual(normalize_course("大数据导论"), "大数据科学导论")
        self.assertEqual(normalize_course("大数据"), "大数据")

    def test_same_major_faculty_teacher_and_course_merge_across_students_and_grades(self):
        data = self.make_workbook([
            [1, "张三", "数学与应用数学", "2301", "数学与统计学院", "尚海锋", "代数与微积分", "23级数学与应用数学同学反映，讲解清楚。"],
            [1, "李四", "数学应用数学", "2401", "数学与统计学院", "尚 海锋老师", "代数与微积分", "24级数学应用数学同学反映，建议增加例题。"],
            [1, "王五", "数学与应用数学", "2401", "数学与统计学院", "尚海锋", "概率论", "24级数学与应用数学同学反映，进度适中。"],
        ])

        result, summary = transform_feedback(data, "第1周.xlsx")

        self.assertEqual(len(result), 2)
        calculus = result[result["课程"] == "代数与微积分"].iloc[0]
        self.assertEqual(calculus["专业"], "数学与应用数学")
        self.assertEqual(calculus["教师"], "尚海锋")
        self.assertIn("讲解清楚", calculus["反馈信息"])
        self.assertIn("建议增加例题", calculus["反馈信息"])
        self.assertEqual(summary["valid_rows"], 3)
        self.assertEqual(summary["merged_rows"], 2)


if __name__ == "__main__":
    unittest.main()

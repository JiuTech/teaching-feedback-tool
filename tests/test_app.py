import io
import unittest

import pandas as pd

from app import (
    normalize_course,
    normalize_major,
    normalize_teacher,
    output_filename,
    transform_feedback,
)


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
        self.assertEqual(normalize_major("数学与应用数学（中外合作办学）"), "数学与应用数学")
        self.assertEqual(normalize_major("数学应用数学(中外合作办学)"), "数学与应用数学")

    def test_course_name_completion_is_conservative(self):
        self.assertEqual(normalize_course("统计软件"), "统计软件课程设计")
        self.assertEqual(normalize_course("大数据导论"), "大数据科学导论")
        self.assertEqual(normalize_course("大数据"), "大数据")
        self.assertEqual(normalize_course("概率论与数理统计"), "概率论与数理统计A2")
        self.assertEqual(normalize_course("概率论与数理统计(含随机过程)"), "概率论与数理统计A2")

    def test_teacher_typo_is_canonicalized(self):
        self.assertEqual(normalize_teacher("尚海峰"), "尚海锋")
        self.assertEqual(normalize_teacher(" 尚 海锋老师 "), "尚海锋")

    def test_output_filename_uses_arabic_week_number(self):
        self.assertEqual(
            output_filename("第十二周原始反馈.xlsx", 1),
            "关于数学与统计学院第12周教学信息反馈.xlsx",
        )
        self.assertEqual(
            output_filename("原始反馈.xlsx", 3),
            "关于数学与统计学院第3周教学信息反馈.xlsx",
        )

    def test_same_grade_major_faculty_teacher_and_course_merge_across_students(self):
        data = self.make_workbook([
            [1, "张三", "数学与应用数学（中外合作办学）", "2401", "数学与统计学院", "尚海锋", "代数与微积分", "24级数学与应用数学（中外合作办学）的同学反映，讲解清楚。"],
            [1, "李四", "数学应用数学", "2402", "数学与统计学院", "尚海峰", "代数与微积分", "24级数学应用数学同学反映，建议增加例题。"],
            [1, "王五", "数学与应用数学", "2401", "数学与统计学院", "尚海锋", "概率论", "24级数学与应用数学同学反映，进度适中。"],
        ])

        result, summary = transform_feedback(data, "第1周.xlsx")

        self.assertEqual(len(result), 2)
        calculus = result[result["课程"] == "代数与微积分"].iloc[0]
        self.assertEqual(calculus["专业"], "数学与应用数学")
        self.assertEqual(calculus["教师"], "尚海锋")
        self.assertIn("讲解清楚", calculus["反馈信息"])
        self.assertIn("建议增加例题", calculus["反馈信息"])
        self.assertEqual(calculus["年级"], "24级")
        self.assertEqual(summary["valid_rows"], 3)
        self.assertEqual(summary["merged_rows"], 2)

    def test_different_grades_remain_separate(self):
        data = self.make_workbook([
            [1, "张三", "数学与应用数学", "2301", "数学与统计学院", "王老师", "概率论", "23级数学与应用数学同学反映，讲解清楚。"],
            [1, "李四", "数学与应用数学", "2401", "数学与统计学院", "王老师", "概率论", "24级数学与应用数学同学反映，建议增加例题。"],
        ])

        result, _ = transform_feedback(data, "第1周.xlsx")

        self.assertEqual(len(result), 2)
        self.assertEqual(set(result["年级"]), {"23级", "24级"})

    def test_repeated_standard_leads_are_removed_once_from_final_sentence(self):
        data = self.make_workbook([
            [3, "张三", "数学与应用数学", "2501", "数学与统计学院", "尚海锋", "概率论", "25级数学与应用数学专业的同学反映，老师在所授的《概率论》课程中，25级数学与应用数学的同学表示，老师在所教授的概率论课程，讲解清晰，逻辑严谨，富有激情。"],
            [3, "李四", "数学与应用数学（中外合作办学）", "2502", "数学与统计学院", "尚海锋", "概率论", "25级数学与应用数学（中外合作办学）的同学反映，希望提前发下课件，有些同学想去预习。"],
        ])

        result, _ = transform_feedback(data, "第3周.xlsx")

        self.assertEqual(len(result), 1)
        feedback = result.iloc[0]["反馈信息"]
        self.assertEqual(
            feedback,
            "25级数学与应用数学专业的同学反映，老师在所授的《概率论》课程中，讲解清晰，逻辑严谨，富有激情，但希望提前发下课件，有些同学想去预习。",
        )
        self.assertEqual(feedback.count("同学反映"), 1)
        self.assertEqual(feedback.count("老师在所授"), 1)


if __name__ == "__main__":
    unittest.main()

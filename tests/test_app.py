import io
import unittest

import pandas as pd

from app import (
    major_from_feedback,
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
        self.assertEqual(normalize_major("统计"), "应用统计学")
        self.assertEqual(normalize_major("数学与应用数学（中外合作办学）"), "数学与应用数学")
        self.assertEqual(normalize_major("数学应用数学(中外合作办学)"), "数学与应用数学")

    def test_course_name_completion_is_conservative(self):
        self.assertEqual(normalize_course("统计软件"), "统计软件课程设计")
        self.assertEqual(normalize_course("大数据导论"), "大数据科学导论")
        self.assertEqual(normalize_course("大数据"), "大数据")
        self.assertEqual(normalize_course("概率论与数理统计"), "概率论与数理统计")
        self.assertEqual(normalize_course("概率论与数理统计(含随机过程)"), "概率论与数理统计A2")

    def test_vague_major_in_feedback_uses_reporter_major(self):
        self.assertEqual(
            major_from_feedback("2025级相关专业的同学反映，课堂节奏适中。", "数学与应用数学"),
            "数学与应用数学",
        )

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

    def test_base_course_is_only_promoted_to_a2_when_a2_is_also_present(self):
        data = self.make_workbook([
            [3, "张三", "会计学", "2501", "数学与统计学院", "王翠萍", "概率论与数理统计A2", "25级会计学专业的同学反映，老师在所授的《概率论与数理统计A2》的课程中，认真细致，条理清晰，但是希望增加练习。"],
            [3, "李四", "会计学", "2505", "数学与统计学院", "王翠萍", "概率论与数理统计", "25级会计学专业的同学反映，老师在所授的《概率论与数理统计》课程中，讲课深入浅出，循序渐进，讲得非常好。"],
        ])

        result, _ = transform_feedback(data, "第3周.xlsx")

        self.assertEqual(len(result), 1)
        self.assertEqual(result.iloc[0]["课程"], "概率论与数理统计A2")
        self.assertEqual(
            result.iloc[0]["反馈信息"],
            "25级会计学专业的同学反映，老师在所授的《概率论与数理统计A2》课程中，讲课深入浅出，循序渐进，认真细致，条理清晰，但是希望增加练习。",
        )

    def test_base_course_stays_unnumbered_without_numbered_source(self):
        data = self.make_workbook([
            [4, "张三", "统计", "2501", "数学与统计学院", "范协铨", "概率论与数理统计", "25级统计专业的同学反映，范老师在所教授的概率论课程中，课上会提问，同学反馈好。"],
        ])

        result, _ = transform_feedback(data, "第4周.xlsx")

        self.assertEqual(result.iloc[0]["专业"], "应用统计学")
        self.assertEqual(result.iloc[0]["课程"], "概率论与数理统计")
        self.assertNotIn("同学反馈好", result.iloc[0]["反馈信息"])
        self.assertEqual(result.iloc[0]["反馈信息"].count("老师在所授"), 1)

    def test_overlapping_praise_is_deduplicated_but_unique_detail_is_kept(self):
        data = self.make_workbook([
            [3, "张三", "数学与应用数学（中外合作办学）", "2501", "数学与统计学院", "尚海锋", "代数与微积分", "25级数学与应用数学专业的同学反映，老师在所授的《代数与微积分》课程中，能够清晰地讲清每一个知识点，板书排版有条理，希望老师能继续用不同颜色的粉笔分清正文和草稿，感谢老师的付出。"],
            [3, "李四", "数学与应用数学", "2503", "数学与统计学院", "尚海锋", "代数与微积分", "25级数学与应用数学专业的同学反映，老师在所授的《代数与微积分》课程中讲的好。"],
            [3, "王五", "数学与应用数学", "2504", "数学与统计学院", "尚海锋", "代数与微积分", "2025级相关专业的同学反映，老师在所授的《代数与微积分》课程中，讲课条理清晰，重难点讲解细致，板书完整，能耐心解答同学们课上提出的数学问题，课堂节奏适中。"],
        ])

        result, _ = transform_feedback(data, "第3周.xlsx")

        self.assertEqual(len(result), 1)
        feedback = result.iloc[0]["反馈信息"]
        self.assertNotIn("讲的好", feedback)
        self.assertEqual(feedback.count("板书"), 1)
        self.assertIn("能耐心解答同学们课上提出的数学问题", feedback)
        self.assertTrue(feedback.endswith("希望老师能继续用不同颜色的粉笔分清正文和草稿，感谢老师的付出。"))


if __name__ == "__main__":
    unittest.main()

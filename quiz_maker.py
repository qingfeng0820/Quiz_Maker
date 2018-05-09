import csv

from docx import Document
from docx.shared import Inches, Pt


class Question:
    def __init__(self, question, answers=None, choices=None):
        self.question = question
        self.choices = choices
        self.answers = self._load_config_answers(answers)

    def answer(self, q_answer):
        self.q_answer = self._get_answers(q_answer)

    def eval_answer(self):
        r = self._do_eval_answer()
        if r == 0:
            print "answer: %s" % self.q_answer
            print "correct answer: %s" % self.answers
        return r

    def _do_eval_answer(self):
        if self.q_answer == self.answers:
            return 1
        return 0

    def _load_config_answers(self, answers):
        return answers

    def _get_answers(self, q_answer):
        return q_answer


class ChoiceQuestion(Question):
    def _do_eval_answer(self):
        if len(self.q_answer) == len(self.answers):
            for ans in self.q_answer:
                if ans not in self.answers:
                    return 0
            return 1
        else:
            return 0

    def _load_config_answers(self, answers):
        config_answers = answers.split("|")
        correct_answers = []
        for answer in config_answers:
            correct_answers.append(chr(65 + int(answer)))
        return correct_answers

    def _get_answers(self, q_answer):
        qu_answer = []
        for i in q_answer:
            qu_answer.append(i)
        return qu_answer


def load(path):
    questions = []
    with open(path, 'rb') as m_file:
        reader = csv.reader(m_file)
        contents = reader
        for row in contents:
            question = ""
            choices = []
            answers = ""
            for i in range(0, len(row)):
                if i == 0:
                    question = row[i]
                if i > 0 and "ANS|" not in row[i]:
                    choices.append(row[i])
                if "ANS|" in row[i]:
                    answers = row[i].split("ANS|")[1]
            if len(choices) > 1:
                questions.append(ChoiceQuestion(question, answers, choices))
            else:
                questions.append(Question(question, answers, choices))
    return questions


def test(questions):
    score = 0
    for q in questions:
        print q.question
        choice = 65
        for c in q.choices:
            print chr(choice) + ") " + c
            choice += 1
        q.answer(raw_input("> "))
        if q.eval_answer() == 1:
            score += 1
        print ""
    print "Score: " + str(score) + "/" + str(len(questions))


def paper(questions, paper_path="test.doc"):
    document = Document()
    for q in questions:
        question_p = document.add_paragraph(q.question, style='List Number')
        if len(q.choices) > 1:
            choice = 65
            for c in q.choices:
                p = document.add_paragraph(chr(choice) + ") " + c, style='List 2')
                paragraph_format = p.paragraph_format
                paragraph_format.line_spacing = Pt(20)
                # paragraph_format.left_indent = Inches(0.3)
                choice += 1
        else:
            document.add_paragraph("", style='List 2')
    document.save(paper_path)


questions = load("q.csv")
print "Questions loaded: " + str(len(questions))

paper(questions)

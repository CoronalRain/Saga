import requests
from lxml import html
import datetime
import inflect
from docx import Document
from docx.shared import Inches

email = "tpk.mindovermath@gmail.com"
password = "Redshift1!"

login_url = "https://mindovermathtutoring.teachworks.com/accounts/login"
participants_url = "https://mindovermathtutoring.teachworks.com/participants"

start_date = datetime.date(2016, 10, 1)
end_date = datetime.date(2016, 10, 30)

valid_names = "A-K"


class Student(object):
    def __init__(self, name, subject, tutor):
        self.name = name
        self.first_name = str.split(name)[0]
        self.last_name = str.split(name)[-1]
        self.subject = subject
        self.tutor = tutor
        self.lessons = []
        self.topics = []

    def __str__(self):
        st = "---------------\n"
        st += self.name + "\n"
        st += self.subject + "\n"
        st += "---------------\n"
        for i in range(len(self.lessons)):
            lesson = self.lessons[i]
            st += "Lesson " + lesson.code + "\n"
            st += lesson.date.isoformat() + "\n"
            st += lesson.notes + "\n"
        st += "Topics: "
        for topic in self.topics:
            st += topic + "; "
        st += "\n\n"

        return st

    def add_lesson(self, lesson):
        self.lessons.append(lesson)

    def add_topics(self, topics):
        self.topics.extend(topics)


class Lesson(object):
    def __init__(self, code, date=datetime.date(9999,12,31), notes="", topics=[]):
        self.code = code
        self.date = date
        self.notes = notes
        self.topics = topics


def main():
    with requests.session() as session_requests:
        # Open the login page.
        response = session_requests.get(login_url)

        # Grab hidden inputs.
        tree = html.fromstring(response.text)
        hidden_inputs = tree.xpath(r"//form//input[@type='hidden']")
        form = {x.attrib["name"]: x.attrib["value"] for x in hidden_inputs}

        # Enter username and password.
        form["user[email]"] = email
        form["user[password]"] = password

        # POST the login request.
        session_requests.post(login_url, data=form)

        # Scape student info.
        students = scrape_info(session_requests)
        write_docx(students)


def scrape_info(session_requests):
    students = []
    student_names = []
    i = 0
    cont = True
    while cont:
        i += 1

        # Visit the next page of lessons.
        response = session_requests.get(participants_url + "?page=" + str(i))
        tree = html.fromstring(response.content)

        # If the last page is reached, break out of the loop.
        if tree.xpath("//div[@id='participants']/div/table/tbody/tr[1]/td[1]/text()")[0] == "No results found.":
            break
        else:
            # Scrape relevant information.
            dates = tree.xpath("//div[@id='participants']/div/table/tbody/tr/td[2]/text()")
            lessons = tree.xpath("//div[@id='participants']/div/table/tbody/tr/td[4]/a/@href")
            subjects = tree.xpath("//div[@id='participants']/div/table/tbody/tr/td[4]/a/text()")
            tutors = tree.xpath("//div[@id='participants']/div/table/tbody/tr/td[5]/text()")
            names = tree.xpath("//div[@id='participants']/div/table/tbody/tr/td[6]/text()")
            statuses = tree.xpath("//div[@id='participants']/div/table/tbody/tr/td[8]/text()")

            for j in range(len(lessons)):
                lesson = str.split(lessons[j], "/")[-1]
                subject = str.split(subjects[j], "-")[-1].strip()
                tutor = tutors[j]
                name = names[j]
                status = statuses[j]

                # Skip this entry if the date is outside of the acceptable range of values.
                date = datetime.date(int(dates[j].split("/")[2]), int(dates[j].split("/")[0]), int(dates[j].split("/")[1]))
                if date < start_date:
                    cont = False
                    break
                if date > end_date:
                    continue

                # Skip this entry if the last name is outside of the acceptable range of values.
                letter = str.upper(name.split()[-1][0])
                if (valid_names == "A-K" and letter > "K") or (valid_names == "L-Z" and letter < "L"):
                    continue

                # Either create a new student or add a lesson to an existing student.
                if status == "Attended":
                    student = None
                    if name not in student_names:
                        student = Student(name, subject, tutor)
                        student_names.append(name)
                        students.append(student)
                    else:
                        for s in students:
                            if s.name == name:
                                student = s
                                break

                    # Visit the lesson page and scrape the internal notes.
                    response = session_requests.get(participants_url + "/" + lesson)
                    tree = html.fromstring(response.content)
                    notes_array = tree.xpath("//div[@class='row participant-notes']/div[2]/span/text()")
                    notes = ""
                    for note in notes_array:
                        notes += note + "\n"

                    # If there are topics listed, parse them out.
                    notes_split = notes.split("Topics:")
                    topics = []
                    if len(notes_split) > 1:
                        notes = notes_split[0]
                        topics = notes_split[-1].split(";")
                        topics = [topic.strip() for topic in topics]
                    student.add_topics(topics)

                    # Add the new lesson for this student.
                    student.add_lesson(Lesson(lesson, date, notes, topics))

    return students


def write_docx(students):
    start = start_date.strftime("%B %d").lstrip("0").replace(" 0", " ")
    end = end_date.strftime("%B %d").lstrip("0").replace(" 0", " ")

    document = Document()

    for i in range(len(students)):
        student = students[i]
        full_name = student.name
        first_name = student.first_name
        subject = student.subject
        tutor = student.tutor
        num = len(student.lessons)
        topics = student.topics

        p = document.add_paragraph("")
        p.add_run(full_name + " - " + subject).underline = True
        p.add_run("\n")
        p.add_run(
            "From " + start + " to " + end + ", " + first_name + " attended " + inflect.engine().number_to_words(num) + " tutoring sessions for " +
            subject + ". " + tutor.split()[0].strip()[0] + " and " + first_name + " were able to cover the following topics:")

        for topic in topics:
            document.add_paragraph(topic, style="List Bullet").paragraph_format.left_indent = Inches(.5)

        document.add_paragraph(
            "Paragraph detailing: (1) improvements, breakthroughs; (2) struggles, solutions; (3) test grades, positive/negative, plan of action; (4) concerns about student; (5) goals for future sessions. Please let us know if you have any questions about " + first_name + "'s progress.\n\n")

        if i % 2 == 1:
            document.add_page_break()

    document.save(start_date.strftime("%B %Y") + " (" + valid_names + ")" + " Auto Progress Reports " +
                  tutor.split()[0].strip()[0] + " " + tutor.split()[-1].strip() + ".docx")

if __name__ == "__main__":
    main()

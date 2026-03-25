from docx import Document
import matplotlib.pyplot as plt
import os


class DocumentTools:
    """
    Toolset for generating DOCX documents
    """

    # 🔥 NEW: Markdown → DOCX formatter
    def add_formatted_paragraph(self, doc, text):
        paragraph = doc.add_paragraph()
        i = 0

        while i < len(text):
            # Bold (**text**)
            if text[i:i+2] == "**":
                i += 2
                start = i
                while i < len(text) and text[i:i+2] != "**":
                    i += 1
                run = paragraph.add_run(text[start:i])
                run.bold = True
                i += 2

            # Italic (*text*)
            elif text[i] == "*":
                i += 1
                start = i
                while i < len(text) and text[i] != "*":
                    i += 1
                run = paragraph.add_run(text[start:i])
                run.italic = True
                i += 1

            else:
                paragraph.add_run(text[i])
                i += 1

    def generate_docx(self, title: str, content: str, output_file: str = "output.docx"):
        """
        Create DOCX with formatted text (bold/italic supported)
        """

        doc = Document()

        # Title
        doc.add_heading(title, 0)

        # Content (UPDATED)
        for line in content.split("\n"):
            if line.strip():
                self.add_formatted_paragraph(doc, line)

        doc.save(output_file)

        return {
            "status": "success",
            "file": output_file
        }

    def generate_docx_with_chart(self, title: str, content: str, data: list, output_file: str = "output.docx"):
        """
        Create DOCX with embedded chart + formatted text
        """

        chart_path = "chart.png"

        # Choose chart type
        chart_type = self.choose_chart_type(data)

        labels = [d[0] for d in data]
        values = [d[1] for d in data]

        plt.figure()

        # 🔥 Use selected chart type
        if chart_type == "pie":
            plt.pie(values, labels=labels, autopct='%1.1f%%')
        else:
            plt.bar(labels, values)

        plt.title("Generated Chart")
        plt.savefig(chart_path)
        plt.close()

        # Create document
        doc = Document()
        doc.add_heading(title, 0)

        for line in content.split("\n"):
            if line.strip():
                self.add_formatted_paragraph(doc, line)

        # Add chart image
        if os.path.exists(chart_path):
            doc.add_paragraph("Chart:")
            doc.add_picture(chart_path)

        doc.save(output_file)

        return {
            "status": "success",
            "file": output_file
        }

    def choose_chart_type(self, data: list):
        """
        Simple logic to choose chart type
        """

        if len(data) <= 5:
            return "pie"
        return "bar"
from docx import Document
import matplotlib.pyplot as plt
import os


class DocumentTools:
    """
    Toolset for generating DOCX documents
    """

    def generate_docx(self, title: str, content: str, output_file: str = "output.docx"):
        """
        Create a basic DOCX file with title + paragraphs
        """

        doc = Document()

        # Title
        doc.add_heading(title, 0)

        # Content
        for line in content.split("\n"):
            if line.strip():
                doc.add_paragraph(line)

        doc.save(output_file)

        return {
            "status": "success",
            "file": output_file
        }

    def generate_docx_with_chart(self, title: str, content: str, data: list, output_file: str = "output.docx"):
        """
        Create DOCX with embedded chart
        data = [("Jan", 10), ("Feb", 20)]
        """

        chart_path = "chart.png"

        # Create chart
        labels = [d[0] for d in data]
        values = [d[1] for d in data]

        plt.figure()
        plt.bar(labels, values)
        plt.title("Generated Chart")
        plt.savefig(chart_path)
        plt.close()

        # Create document
        doc = Document()
        doc.add_heading(title, 0)

        for line in content.split("\n"):
            if line.strip():
                doc.add_paragraph(line)

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

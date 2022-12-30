import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QRadioButton, QPushButton, QProgressBar, QFileDialog, QMessageBox, QCheckBox, QListWidget
import facebook
import requests
import re
import docx

class CommentExtractor(QMainWindow):
    def __init__(self):
        super().__init__()

        # Set window title and size
        self.setWindowTitle("Facebook Comment Extractor")
        self.resize(500, 400)

        # Create input field for post ID
        self.post_id_label = QLabel(self)
        self.post_id_label.setText("Post ID(s):")
        self.post_id_label.move(20, 20)
        self.post_id_entry = QLineEdit(self)
        self.post_id_entry.move(20, 50)

        # Create checkbox for single/multiple post selection
        self.multiple_posts_checkbox = QCheckBox(self)
        self.multiple_posts_checkbox.setText("Extract comments from multiple posts")
        self.multiple_posts_checkbox.move(20, 90)

        # Create radio buttons for output format
        self.format_label = QLabel(self)
        self.format_label.setText("Output format:")
        self.format_label.move(20, 130)
        self.word_button = QRadioButton(self)
        self.word_button.setText("Word")
        self.word_button.move(20, 160)
        self.pdf_button = QRadioButton(self)
        self.pdf_button.setText("PDF")
        self.pdf_button.move(20, 190)

        # Create checkbox for single/multiple file selection
        self.multiple_files_checkbox = QCheckBox(self)
        self.multiple_files_checkbox.setText("Create a separate file for each post")
        self.multiple_files_checkbox.move(20, 230)

        # Create output file selection button
        self.output_button = QPushButton(self)
        self.output_button.setText("Select output file")
        self.output_button.move(20, 270)
        self.output_button.clicked.connect(self.select_output)

        # Create list widget to display comments
        self.comments_list = QListWidget(self)
        self.comments_list.move(300, 20)

        # Create extract button
        self.extract_button = QPushButton(self)
        self.extract_button.setText("Extract comments")
        self.extract_button.move(20, 300)
        self.extract_button.clicked.connect(self.extract)
        self.progress_bar = QProgressBar(self)
        self.progress_bar.move(20, 340)
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)

    def select_output(self):
        """
        Opens a file selection dialog for the user to choose the output file.
        """
        self.output_file, _ = QFileDialog.getSaveFileName(self, "Select output file", "", "Word Document (*.docx);;PDF (*.pdf)")

    def extract(self):
        """
        Extracts comments from the specified Facebook post(s) and creates a Word or PDF document with the extracted comments.
        """
        # Get post ID(s) and output format
        post_ids = self.post_id_entry.text().split(",")
        output_format = self.word_button.isChecked()

        # Validate input
        if not post_ids:
            # Display error message if no post ID(s) are entered
            QMessageBox.warning(self, "Error", "Please enter a post ID.")
            return
        if not self.output_file:
            # Display error message if no output file is selected
            QMessageBox.warning(self, "Error", "Please select an output file.")
            return

        # Extract comments from Facebook posts
        comments = []
        for post_id in post_ids:
            try:
                comments += self.extract_comments(post_id)
            except Exception as e:
                # Display error message if an exception occurs
                QMessageBox.warning(self, "Error", f"An error occurred while extracting comments from post {post_id}: {e}")
                return

        # Clear list widget and add extracted comments
        self.comments_list.clear()
        for comment in comments:
            self.comments_list.addItem(comment["text"])

        # Create output file(s)
        if output_format:
            # Create Word document
            doc = template_doc
            for i, comment in enumerate(comments):
                doc.add_paragraph(f"Comment {i+1}: {comment['text']}")
            doc.save(self.output_file)
        else:
            # Create PDF document (TODO)
            pass

        # Display success message
        QMessageBox.information(self, "Success", "Comments successfully extracted.")

    def extract_comments(self, post_id):
        """
        Extracts all comments from a Facebook post.
        """
        # Set up Facebook Graph API client
        token = "YOUR_TOKEN_HERE"
        graph = facebook.GraphAPI(access_token=token, version="3.0")

        # Get post data
        post = graph.get_object(id=post_id)

        # Extract comments
        comments = []
        comments_data = graph.get_connections(post["id"], "comments")
        while True:
            try:
                # Get comments for current page
                for comment in comments_data["data"]:
                    comments.append(comment)

                # Get next page of comments
                comments_data = requests.get(comments_data["paging"]["next"]).json()
            except KeyError:
                # No more pages of comments
                break

        return comments

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = CommentExtractor()
    window.show()
    sys.exit(app.exec_())



import os
import shutil
import pygit2
import subprocess
import pypandoc
import convert_doc_controller
from shutil import copy, move
from tkinter import filedialog, simpledialog, ttk, messagebox, Frame, TOP, BOTTOM
from tkinter import *
from resolutions import horizontal_resolution, vertical_resolution

class convert_doc_gui:
    def __init__(self, master):

        self.master = master
        master.title("ConvertDoc")
        self.label = Label(master, text="ConvertDoc")
        self.label.pack()
        top_frame = Frame(master)
        top_frame.pack(side=TOP)
        bottom_frame = Frame(master)
        bottom_frame.pack(side=BOTTOM)
        west_frame = Frame(master)
        west_frame.pack(side=LEFT)
        east_frame = Frame(master)
        east_frame.pack(side=RIGHT)

        self.progress_bar = ttk.Progressbar(bottom_frame, orient='horizontal', mode='determinate', maximum=100)
        self.convert_to_md_button = Button(master, text="Convert DOCX to MD", command=lambda : convert_doc_controller.convert_docx_to_markdown(self.progress_bar))
        self.convert_to_md_button.pack(padx=10, pady=10, expand=True, side=LEFT)
        self.upload_to_git_button = Button(master, text="Upload DOCX to Git Repo", command=lambda : self.upload_to_git_repo())
        self.upload_to_git_button.pack(padx=10, pady=10, expand=True, side=LEFT)
        self.convert_to_docx_button = Button(master, text="Convert MD to DOCX", command=lambda : convert_doc_controller.convert_markdown_to_docx_button(self.progress_bar))
        self.convert_to_docx_button.pack(padx=10, pady=10, expand=True, side=LEFT)
        self.progress_bar.pack(fill=BOTH, padx=10, pady=10, anchor='center', expand=True)


    def upload_to_git_repo(self):

        # Determine if the user needs to clone
        has_repo = messagebox.askyesno(title="Repo Exists?", message="Do you have an existing Git repo you want to use?")

        if not has_repo:
            # Get the Repository Url
            url = simpledialog.askstring("Git Repo Url", "Enter the Git Repository URL to clone from")
            if url == None:
                return

            elif url == "":
                return

            # Get the desired location to clone the repository to.
            messagebox.showinfo("Folder to clone to", "Select a folder to clone the git repo to.")
            repo_location = filedialog.askdirectory()

            if repo_location == None:
                return

            elif repo_location == "":
                return

            # Get the desired branch to clone from
            branch = simpledialog.askstring("Git Branch", "Enter the name of the branch to clone from")

            if branch == None:
                return

            elif branch == "":
                return

            # Attempt to clone
            try:
                pygit2.clone_repository(url=url, path=repo_location, checkout_branch=branch)

            except pygit2.GitError:
                messagebox.showwarning("Git Repository", "Git Repository not found.")
                return
            except ValueError:
                messagebox.showwarning(title="Folder is not empty", message="Selected folder is not empty.")
                return
        elif has_repo:
            # Get the location of the repository
            messagebox.showinfo("Repo Location", "Select a folder where your git repo exists.")
            repo_location = filedialog.askdirectory()
            if repo_location == None:
                return
            elif repo_location == "":
                return

            # Get the desired branch to work on
            branch = simpledialog.askstring(title="Git Branch", prompt="Enter the name of the branch to work on.")

            if branch == None:
                return
            elif branch == "":
                return

            try:
                repo = pygit2.Repository(path=repo_location)

            except pygit2.GitError:
                messagebox.showerror(title="Git Error", message="This folder does not contain an existing git repository.")
                return

            # Get the status of the current branch
            status = repo.status()
            if bool(status):
                messagebox.showerror(title="Uncommitted changes",
                                     message="There are uncommitted changes to this repo. Please resolve them.")
                return

            # Checkout the branch requested
            os.chdir(repo_location)
            try:
                subprocess.run(["git.exe", "checkout", f"{branch}"], shell=True, check=True)

            # Pull the repository
                subprocess.run(["git.exe", "pull"], shell=True, check=True)

            except subprocess.SubprocessError:
                messagebox.showerror(title="Error", message="Could not find git protocol. Install Git first")
                return

        # Convert the doc
        document = self.convert_docx_to_markdown()

        if document is None:
            return

        # Update Progress
        self.progress_bar["value"] = 0
        self.progress_bar["value"] += 20

        # Configure Git
        try:
            # Git Add
            subprocess.run(["git.exe", r"add", "."], shell=True, check=True)
            self.progress_bar["value"] += 20

            # Git Commit
            commit_message = simpledialog.askstring(title="Commit message?", prompt="Enter a commit message.")
            if commit_message is None:
                return

            elif commit_message == "":
                return

            subprocess.run(["git.exe", "commit", "-m", commit_message], shell=True, check=True)
            self.progress_bar["value"] += 20

            # Git push
            subprocess.run(["git.exe", "push"], shell=True, check=True)

            self.progress_bar["value"] += 20

        except subprocess.SubprocessError:
            messagebox.showerror(title="Error", message="Could not find git protocol. Install Git first")
            return

        path_away_from_repo = r"C:\\"
        os.chdir(path_away_from_repo)
        print("Done")

        self.progress_bar["value"] += 20

def remove_folder_recursive(folder):
    for root, dirs, files in os.walk(folder, topdown=False):
        for name in files:
            os.remove(os.path.join(root, name))
        for name in dirs:
            os.rmdir(os.path.join(root, name))
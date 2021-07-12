import os
import shutil
import pygit2
import subprocess
import pypandoc
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

        self.convert_to_md_button = Button(master, text="Convert DOCX to MD", command=lambda : self.convert_docx_to_markdown())
        self.convert_to_md_button.pack(padx=10, pady=10, expand=True, side=LEFT)

        self.upload_to_git_button = Button(master, text="Upload DOCX to Git Repo", command=lambda : self.upload_to_git_repo())
        self.upload_to_git_button.pack(padx=10, pady=10, expand=True, side=LEFT)

        self.convert_to_docx_button = Button(master, text="Convert MD to DOCX", command=lambda : self.convert_markdown_to_docx())
        self.convert_to_docx_button.pack(padx=10, pady=10, expand=True, side=LEFT)

        self.progress_bar = ttk.Progressbar(bottom_frame, orient='horizontal', mode='determinate', maximum=100)
        self.progress_bar.pack(fill=BOTH, padx=10, pady=10, anchor='center', expand=True)

    def convert_markdown_to_docx(self):

        # Initialize Progress Bar Value
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = 100

        # Select a Markdown File
        input_filename = filedialog.askopenfilename(title="Select Markdown File", filetypes=(
            ("Markdown files", "*.md"), ("all files", "*.*")))

        if input_filename is '':
            return

        # Wrap both the input file and the output file for use with pandoc
        # Setup pandoc args
        input_filename = input_filename.replace("/", "\\")

        output_filename = input_filename.replace('.md', '.docx')

        if output_filename is '':
            return

        # Update Progress
        self.progress_bar["value"] += 10

        # First to html. Then later to docx
        # Ask the user where their images are
        has_images = messagebox.askquestion("Images", "Does the markdown document have local images associated"
                                                      "with it?")
        img_dir = None
        html_tempfile = None
        convert_to_html='html'
        default_encoding='utf-8'
        extracted_media = None
        if has_images == 'yes':
            img_dir_options = {}
            img_dir_options['title'] = "Select the image folder"
            messagebox.showinfo("Image Directory Location", "Select the directory where the associated markdown images exist"
                                                        "in reference to the markdown document.")
            img_dir = filedialog.askdirectory()
            html_tempfile = output_filename + "_.html"
            extracted_media = [f'--extract-media={img_dir}']

        elif has_images == 'no':
            html_tempfile = output_filename + "_.html"


        # Change working directory to that of the input file
        output_file_path = os.path.abspath(output_filename)
        working_directory = os.path.dirname(output_file_path)
        os.chdir(working_directory)

        # Update Progress
        self.progress_bar["value"] += 30

        # Execute Pandoc command
        try:
            if extracted_media is None:

                html_convert = pypandoc.convert_file(source_file=input_filename, to=convert_to_html,
                                                     outputfile=html_tempfile, encoding=default_encoding)
            else:

                html_convert = pypandoc.convert_file(source_file=input_filename, to=convert_to_html,
                                                 outputfile=html_tempfile, encoding=default_encoding,
                                                 extra_args=extracted_media)
        except RuntimeError as run_time_error:
            raise run_time_error
        except OSError as os_error:
            raise os_error
        print("Pandoc command executed.")
        convert_to_docx='docx'
        default_encoding = 'utf-8'
        if has_images == 'yes':

            if not output_filename.endswith('.docx'):
                print("Proper filename not supplied, appending extension")
                output_filename += '.docx'
            extracted_media = [f'--extract-media={img_dir}']

        elif has_images == 'no':

            if not output_filename.endswith('.docx'):
                print("Proper filename not supplied, appending extension")
                output_filename += '.docx'
            extracted_media = None

        # Update Progress
        self.progress_bar["value"] += 10

        # Execute Pandoc command
        try:

            if extracted_media is None:

                docx_convert = pypandoc.convert_file(source_file=html_tempfile, to=convert_to_docx,
                                                     outputfile=output_filename, encoding=default_encoding)
            else:
                docx_convert = pypandoc.convert_file(source_file=html_tempfile, to=convert_to_docx,
                                                     outputfile=output_filename, encoding=default_encoding,
                                                     extra_args=extracted_media)
        except RuntimeError as run_time_error:
            raise run_time_error
        except OSError as os_error:
            raise os_error
        print("Pandoc command executed.")

        # Update Progress
        self.progress_bar["value"] += 40

        # Delete Html tempfile
        os.remove(html_tempfile)
        print("Deleted temporary html file")

        # Ask if the save file should be saved in a different location than the original
        save_in_different_location = messagebox.askquestion("Save Location", "Save the word document in a different"
                                                                             "location from the original?")
        if save_in_different_location == 'yes':
            # Get save file name, then copy the file over to that location as that name
            save_filename = filedialog.asksaveasfilename(title="Save Word File", filetypes=(
                ("Word files", "*.docx"), ("all files", "*.*")))

            try:
                if save_filename.endswith('.docx'):
                    copy(output_filename, save_filename)
                else:
                    print("Proper filename not supplied, appending extension.")
                    save_filename += '.docx'
                    copy(output_filename, save_filename)
            except shutil.SameFileError:
               messagebox.showwarning("File Exists", f"The file {save_filename} already exists.")
               os.remove(output_filename)
               self.progress_bar["value"] = 0
               return

            # Delete output file
            os.remove(output_filename)

            # Update Progress
            self.progress_bar["value"] += 10

            # Inform the user that the process is complete
            messagebox.showinfo("Process Complete", f"Word file is in {save_filename}")
        else:
            # Update Progress
            self.progress_bar["value"] += 10

            # Inform the user that the process is complete
            messagebox.showinfo("Process Complete", f"Word file is in {output_filename}")

    def convert_docx_to_markdown(self):

        # Initialize Progress Bar Value
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = 100

        # Select a Word Document
        input_filename = filedialog.askopenfilename(title="Select Word Document", filetypes=(
            ("Word files", "*.docx"),("all files", "*.*")))

        if input_filename is '':
            return
        elif input_filename == None:
            return

        # Prep the input file, setup working directory
        # Working Directory should be the one where the Input File is selected
        input_filename_without_path = os.path.basename(input_filename)
        input_filename_without_spaces = input_filename_without_path.replace(" ", "_")
        input_filename_without_spaces_or_extension = input_filename_without_spaces.replace(".docx", "")

        attachments_folder = input_filename_without_spaces_or_extension
        attachments_folder = attachments_folder.replace("\\", "/")

        print(attachments_folder)

        # Wrap both the input file and the output file for use with pandoc
        # Setup pandoc args
        input_filename = input_filename.replace("/", "\\")

        # Pick a save location instead, then append tha input_filename to that location for pandoc's output

        output_filename = filedialog.asksaveasfilename(title="Save Markdown File", filetypes=(
            ("Markdown files", "*.md"), ("all files", "*.*")))

        if output_filename is '':
            return
        elif output_filename is None:
            return
        convert_md = 'md'
        default_encoding = 'utf-8'
        extracted_media = None
        if output_filename.endswith('.md'):
            extracted_media = [f"--extract-media={attachments_folder}"]
        else:
            print("Proper filename not supplied, appending extension")
            output_filename += '.md'
            extracted_media = [f"--extract-media={attachments_folder}"]

        # Update Progress
        self.progress_bar["value"] += 10

        # Change working directory to that of the input file
        output_file_path = os.path.abspath(output_filename)
        working_directory = os.path.dirname(output_file_path)
        os.chdir(working_directory)

        # Update Progress
        self.progress_bar["value"] += 10

        # Update Progress
        self.progress_bar["value"] += 30

        # Execute Pandoc command
        try:
            convert_to_md = pypandoc.convert_file(source_file=input_filename, to=convert_md,
                                                  encoding=default_encoding, outputfile=output_filename,
                                                  extra_args=extracted_media)
        except RuntimeError as run_time_error:
            raise run_time_error
        except OSError as os_error:
            raise os_error

        print("Pandoc command executed.")

        # Update Progress
        self.progress_bar["value"] += 40

        # Update Progress
        self.progress_bar["value"] += 10

        return output_filename

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


def main():
    root = Tk()
    gui = convert_doc_gui(root)
    minwidth = 854
    minheight = 480

    screenwidth = minwidth
    screenheight = minheight

    fullscreen_width = root.winfo_screenwidth()
    fullscreen_height = root.winfo_screenheight()

    if fullscreen_width == horizontal_resolution.UHD.value and fullscreen_height == vertical_resolution.UHD.value:
        screenwidth = horizontal_resolution.QVGA.value
        screenheight = vertical_resolution.QVGA.value

    if fullscreen_width < horizontal_resolution.QHD.value and fullscreen_height < vertical_resolution.QHD.value:
        screenwidth = horizontal_resolution.VGA.value
        screenheight = vertical_resolution.VGA.value


    root.geometry('%sx%s' % (screenwidth, screenheight))
    root.mainloop()

if __name__ == '__main__':
    main()
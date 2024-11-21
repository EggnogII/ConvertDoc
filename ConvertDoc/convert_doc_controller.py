import os
import pypandoc
from tkinter import filedialog, simpledialog, ttk, messagebox

def document_selector(selection_title, filetype):
    filename = filedialog.askopenfilename(title=selection_title, filetypes=
    (("Documents", f"*.{filetype}"), ("all files", "*.*")))

    return filename


def convert_markdown_to_docx_button(progress_bar):
        # Initialize Progress Bar Value
        progress_bar["value"] = 0
        progress_bar["maximum"] = 100

        # Select a Markdown File
        input_filename = document_selector(selection_title="Select Markdown File", filetype="md")

        if input_filename == '':
            return

        # Wrap both the input file and the output file for use with pandoc
        # Setup pandoc args
        input_filename = input_filename.replace("/", "\\")
        output_filename = input_filename.replace('.md', '.docx')

        if output_filename == '':
            return

        # Update Progress
        progress_bar["value"] += 10

        # First to html. Then later to docx
        # # Ask the user where their images are
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
        progress_bar["value"] += 30

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
        progress_bar["value"] += 10
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
        progress_bar["value"] += 40
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
               progress_bar["value"] = 0
               return

            # Delete output file
            os.remove(output_filename)

            # Update Progress
            progress_bar["value"] += 10
            # Inform the user that the process is complete

            messagebox.showinfo("Process Complete", f"Word file is in {save_filename}")

        else:
            # Update Progress
            progress_bar["value"] += 10

            # Inform the user that the process is complete
            messagebox.showinfo("Process Complete", f"Word file is in {output_filename}")

def convert_docx_to_markdown(progress_bar):

        # Initialize Progress Bar Value
        progress_bar["value"] = 0
        progress_bar["maximum"] = 100

        # Select a Word Document
        input_filename = document_selector("Select a Word Document", "docx")

        if input_filename == '':
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

        if output_filename == '':
            return

        elif output_filename == None:
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
        progress_bar["value"] += 10

        # Change working directory to that of the input file
        output_file_path = os.path.abspath(output_filename)
        working_directory = os.path.dirname(output_file_path)
        os.chdir(working_directory)

        # Update Progress
        progress_bar["value"] += 10

        # Update Progress
        progress_bar["value"] += 30

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
        progress_bar["value"] += 40

        # Update Progress
        progress_bar["value"] += 10
        return output_filename
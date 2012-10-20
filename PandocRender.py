import sublime, sublime_plugin
import webbrowser
import tempfile
import os
import re
import sys
import subprocess

plugin_path = os.path.dirname(os.path.abspath(__file__))

class PandocRenderCommand(sublime_plugin.TextCommand):
    """ render file contents to HTML and, optionally, open in your web browser"""

    def getTemplatePath(self, filename):
        path = os.path.join(plugin_path, filename)
        if not os.path.isfile(path):
            raise Exception(filename + " file not found!")
        return path

    def is_enabled(self):
        return self.view.score_selector(0, "text.html.markdown") > 0

    def is_visible(self):
        return True

    def run(self, edit, target="html", openAfter=True, writeBeside=False, commandArgs=[]):

        # grab contents of buffer
        region = sublime.Region(0, self.view.size())
        encoding = self.view.encoding()
        if encoding == 'Undefined':
            encoding = 'UTF-8'
        elif encoding == 'Western (Windows 1252)':
            encoding = 'windows-1252'
        contents = self.view.substr(region)

        # write buffer to temporary file
        tmp_md = tempfile.NamedTemporaryFile(delete=False, suffix=".md")
        tmp_md.write(contents)
        tmp_md.close()

        # output file...
        suffix = "." + target
        if writeBeside:
            output_filename = os.path.splitext(self.view.file_name())[0]+suffix
            if not self.view.file_name(): raise Exception("Buffer must be saved before 'writeBeside' can be used.")
        else:
            tmp_html = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            tmp_html.close()
            output_filename=tmp_html.name

        # build output
        cmd = ['pandoc']
        cmd.append('--data-dir={0}'.format(plugin_path));
        cmd += commandArgs

        # Extra arguments to pass into pandoc, e.g.:
        #     <!-- [[ PANDOC --smart --no-wrap ]] -->
        matches = re.finditer(r'<!--\s+\[\[ PANDOC (?P<args>.*) \]\]\s+-->', contents)
        for match in matches:
            cmd += match.groupdict()['args'].split(' ')

        # Destination file
        cmd.append("-o")
        cmd.append(output_filename)

        # Source file
        cmd.append(tmp_md.name)

        try:
            print('Executing command: ' + ' '.join(cmd))
            output = subprocess.Popen(cmd, stdout = subprocess.PIPE).communicate()
            print(output)
        except Exception as e:
            sublime.error_message("Unable to execute Pandoc.  \n\nDetails: {0}".format(e))

        print "Wrote:", output_filename

        if openAfter:
            if target == "html":
                webbrowser.open_new_tab(output_filename)
            # perhaps there is a better way of handling the DocX opening...?
            elif target == "docx" and sys.platform == "win32":
                os.startfile(output_filename)
            elif target == "docx" and sys.platform == "mac":
                subprocess.call(["open", output_filename])
            elif target == "docx" and sys.platform == "posix":
                subprocess.call(["xdg-open", output_filename])
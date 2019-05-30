import configparser
from datetime import datetime
import docx
import os


class File:
    """Class that handles files."""

    def __init__(self, fname, config_fname='basic.ini'):
        self.fname = fname
        self.format = Format(config_fname)
        self.paragraphs = []

    def __getitem__(self, i):
        return self.paragraphs[i]

    def get_modified_text(self):
        """Return a list of div tags and paragraphs wrapped in p tags."""
        result = [self.format.div_tag]
        for par in self.paragraphs:
            string = par.get_string()
            if string:
                wr_par = self.format.p_tag.replace('justify', par.align) + \
                         string + self.format.p_end_tag
                result.append(wr_par)
            else:
                result.append(self.format.div_end_tag + self.format.div_tag)
        result.append(self.format.div_end_tag)

        return result


class TxtFile(File):
    """Subclass for .txt documents."""

    def __init__(self, fname):
        super().__init__(fname)
        with open(self.fname, 'r', encoding='cp1251') as f:
            for par in f.readlines():
                self.paragraphs.append(TxtPar(par))

    def __str__(self):
        return ''.join([par.get_text for par in self.paragraphs])


class DocxFile(File):
    """Subclass for .docx documents."""

    def __init__(self, fname):
        super().__init__(fname)
        for par in docx.Document(self.fname).paragraphs:
            self.paragraphs.append(DocxPar(par))

    def __str__(self):
        return '\n'.join(list(map(str, self.paragraphs)))


class Par:
    """Class for paragraphs of text."""

    def __init__(self, par):
        self.paragraph = par
        self.align = 'justify'


class TxtPar(Par):
    """Subclass for paragraphs of text from .txt documents."""

    def __str__(self):
        return self.paragraph

    def get_string(self):
        return self.paragraph.strip()


class DocxPar(Par):
    """Subclass for paragraphs of text from .docx documents."""

    def __init__(self, par):
        super().__init__(par)
        if 'RIGHT' in str(self.paragraph.alignment):
            self.align = 'right'

    def __str__(self):
        return self.paragraph.text

    def get_string(self):
        """Return the paragraph with basic formatting added."""
        runs = self.paragraph.runs
        pairs_of_runs = list(zip(runs[:], runs[1:]))
        if runs:
            cur = runs[0]
            string = DocxPar.add_tags(cur) + cur.text

            for pair_of_runs in pairs_of_runs:
                prev, cur = pair_of_runs
                tags = DocxPar.add_tags(cur, prev)
                string += tags + cur.text

            string += DocxPar.add_tags(None, cur)
            return string

        return ''

    @staticmethod
    def add_tags(cur, prev=None):
        """Compare formatting of two runs and return proper html tags."""
        tags = ''
        if cur:
            if cur.bold and (not prev or not prev.bold):
                tags += '<b>'
            if cur.italic and (not prev or not prev.italic):
                tags += '<i>'
            if cur.underline and (not prev or not prev.underline):
                tags += '<u>'
        if prev:
            if (not cur or not cur.underline) and prev.underline:
                tags += '</u>'
            if (not cur or not cur.italic) and prev.italic:
                tags += '</i>'
            if (not cur or not cur.bold) and prev.bold:
                tags += '</b>'
        return tags


class Format:
    """Class that reads text formatting settings from the config file."""

    def __init__(self, config_fname):
        config = configparser.ConfigParser()
        config.read(config_fname)
        section = config['DEFAULT']
        ind = '{0}{1};'.format(section['text-indent'],
                               section['text-indent-units'])
        w = '{0}{1};'.format(section['width'], section['width-units'])
        p_mar = Format.get_margins('p', section)
        div_mar = Format.get_margins('div', section)

        self.div_tag = f"<div style='{div_mar} text-indent: {ind} width: {w}'>"
        self.p_tag = f"<p align='justify' style='{p_mar}'>"
        self.div_end_tag = '</div>'
        self.p_end_tag = '</p>'

    @staticmethod
    def get_margins(tag, section):
        top = section[f'{tag}-margin-top']
        right = section[f'{tag}-margin-right']
        bottom = section[f'{tag}-margin-bottom']
        left = section[f'{tag}-margin-left']
        units = section[f'{tag}-margin-units']

        if not all([top, right, bottom, left]):
            return 'margin: 0;'
        elif top == right == bottom == left:
            return 'margin: {1}{0};'.format(units, top)
        elif top == bottom and right == left:
            return 'margin: {1}{0} {2}{0};'.format(units, top, right)
        elif top != bottom and right == left:
            return 'margin: {1}{0} {2}{0} {3}{0};'.format(units, top,
                                                          right, bottom)
        else:
            return 'margin: {1}{0} {2}{0} {3}{0} {4}{0};'.format(units,
                                                                 top, right,
                                                                 bottom, left)


if __name__ == "__main__":

    timestamp = datetime.today().strftime('_%d_%m_%Y')
    idir = os.path.abspath(os.getcwd())
    odir = os.path.abspath(os.path.join(idir, 'Converted' + timestamp))

    if not os.path.exists(odir):
        os.mkdir(odir)

    for fname in os.listdir(idir):

        file_name, file_format = os.path.splitext(fname)
        if file_format == '.txt':
            document = TxtFile(fname)
        elif file_format in ['.docx', '.doc']:
            document = DocxFile(fname)
        else:
            continue

        html_document = document.get_modified_text()

        fpath = os.path.join(odir, file_name + '.txt')
        with open(fpath, 'w', encoding='cp1251') as f:
            for line in html_document:
                f.write(line)

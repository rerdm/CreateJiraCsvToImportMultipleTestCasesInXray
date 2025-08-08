import codecs


class ChangeCsvCodingFromAsiToUTF8:
    """
    In this class the csv file is passed, which was adapted for JIRa.
    In order for all umlauts to be displayed correctly in JIRA, the file must conform to the UTF-8 standard.
    The class changes the encoding of the file so that a JIRA import does not show <?> in the text.
    """

    def __init__(self, csv_name):
        self.csv_name = csv_name

    def convert_encoding(self):
        # Actual coding
        aktuelle_codierung = 'cp1252'  # Example: ANSI (Windows-1252)

        with codecs.open(self.csv_name, 'r', encoding=aktuelle_codierung) as datei:
            inhalt = datei.read()

        # Open file and save it as UTF-8-coded file
        with codecs.open(self.csv_name, 'w', encoding='utf-8') as datei:
            datei.write(inhalt)

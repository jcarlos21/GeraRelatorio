from docx import Document


document = Document()

class AnalysisFunc:
    
    def __init__ (self, document):
        self.document = document
    
    def rR1 (self, potBefore, potAfter):
        if (float(potAfter.strip()) + 1) >= float(potBefore.strip()):
            return 'OK'
        else:
            return 'X'

    def rR2 (self, potMedia, potAfter):
        if (float(potAfter.strip()) + 2.5) >= float(potMedia.strip()) :
            return 'OK'
        else:
            return 'X'

    def rR3 (self, potAfter):
        if float(potAfter.strip()) > -21.00:
            return 'OK'
        else:
            return 'X'

    def rStatus (self, rR1, rR2, rR3):
        if rR1 == rR2 == rR3 == 'OK':
            return 'APROVADO'
        else:
            return 'REAVALIAR'

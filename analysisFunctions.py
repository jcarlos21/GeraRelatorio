from docx import Document


document = Document()

class AnalysisFunc:
    
    def __init__ (self, document):
        self.document = document
    
    def rR1 (self, potBefore, potAfter):
        if (potAfter + 1) >= potBefore:
            return 'OK'
        else:
            return 'X'

    def rR2 (self, potMedia, potAfter):
        if (potAfter + 2.5) >= potMedia :
            return 'OK'
        else:
            return 'X'

    def rR3 (self, potAfter):
        if potAfter > -21.00:
            return 'OK'
        else:
            return 'X'

    def rStatus (self, rR1, rR2, rR3):
        if rR1 == rR2 == rR3 == 'OK':
            return 'APROVADO'
        else:
            return 'REAVALIAR'
    
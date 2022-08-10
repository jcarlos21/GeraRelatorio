from docx import Document


document = Document()

class AnalysisFunc:
    
    def __init__ (self, document):
        self.document = document
    
    def requerimentsR1 (self, potBefore, potAfter):
        if (potAfter - 1) >= potBefore:
            return 'OK'
        else:
            return 'X'

    def requerimentsR2 (self, potMedia, potBefore):
        if (potMedia - 1) >= potBefore:
            return 'OK'
        else:
            return 'X'

    def requerimentsR3 (self, potAfter):
        if potAfter > -21.00:
            return 'OK'
        else:
            return 'X'

    def requerimentsPot (self, requerimentsR1, requerimentsR2, requerimentsR3):
        if requerimentsR1 == requerimentsR2 == requerimentsR3 == 'OK':
            return 'APROVADO'
        else:
            return 'REAVALIAR'
    
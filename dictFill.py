

class WriteDict:

	def __init__(self, dadosDict={'dado': 1}, celula='', entidade='', endereco='', potMedia='', potBefore='', potAfter=''):
		self.dadosDict = dadosDict
		self.celula = celula
		self.entidade = entidade
		self.endereco = endereco
		self.potMedia = potMedia
		self.potBefore = potBefore
		self.potAfter = potAfter

	def fillDict (self):
		if not self.dadosDict.get(self.celula):
			self.dadosDict[self.celula] = {self.entidade: [self.endereco, self.potMedia, self.potBefore, self.potAfter]}
			# return self.dadosDict
			retorno = "Escola adicionada! Não esqueça de adicionar a imagem na pasta /img com o a legenda '01.JPG', '02.JPG', etc."
			return retorno
		else:
			# OBS: Note que se a entidade já existir, o valor dela será substituído por [endereco, potMedia, potBefore, potAfter]
			self.dadosDict[self.celula][self.entidade] = [self.endereco, self.potMedia, self.potBefore, self.potAfter]
			# return self.dadosDict
			retorno = "Escola adicionada! Não esqueça de adicionar a imagem na pasta /img com o a legenda '01.JPG', '02.JPG', etc."
			return retorno


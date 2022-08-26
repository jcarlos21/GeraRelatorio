# ============================================= Dados Coletados =================================================== #

"""
Atributos das entidades:
    - Endereço
	- Trecho (preparado pelo próprio software)
	- Range considerado
	- Potência antes da corretiva
	- Potência depois da corretiva
	- Média de potência atual
	- Imagem referente ao range (previamente salva pelo usuário na pasta do software)
"""

# celula = ''
# entidade = ''
# endereco = ''
# potMedia = 0
# potBefore = 0
# potAfter = 0

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
			retorno = 'Escola adicionada!'
			return retorno
		else:
			# OBS: Note que se a entidade já existir, o valor dela será substituído por [endereco, potMedia, potBefore, potAfter]
			self.dadosDict[self.celula][self.entidade] = [self.endereco, self.potMedia, self.potBefore, self.potAfter]
			# return self.dadosDict
			retorno = 'Escola adicionada!'
			return retorno


# teste = WriteDict()

# dadosGerais = teste.fillDict()

# dicionario = {'Carlos': {'Hobby': ['Tocar', 'Cantar']}}
# dicionario['Carlos']
# {'Hobby': ['Tocar', 'Cantar']}
# dicionario['Carlos']['Hobby']
# ['Tocar', 'Cantar']
# dicionario['Carlos']['Hobby'][0]
# 'Tocar'

# dicionario = {'Carlos': {'Hobby': ['Tocar', 'Cantar']}}
# dicionario['Carlos']['Graduação'] = ['C&T', 'Eng. Telecomunicações']
# dicionario
# {'Carlos': {'Hobby': ['Tocar', 'Cantar'], 'Graduação': ['C&T', 'Eng. Telecomunicações']}}

# {'Carlos': {'Hobby': ['Tocar', 'Cantar'], 'Graduação': ['C&T', 'Eng. Telecomunicações']}}

# DadosGerais = {celula: {entidade: [endereco, potMedia, potBefore, potAfter],
# 						entidade: [endereco, potMedia, potBefore, potAfter],
# 						entidade: [endereco, potMedia, potBefore, potAfter]},
				
# 				celula: {entidade: [endereco, potMedia, potBefore, potAfter],
# 						entidade: [endereco, potMedia, potBefore, potAfter],
# 						entidade: [endereco, potMedia, potBefore, potAfter]},
				
# 				celula: {entidade: [endereco, potMedia, potBefore, potAfter],
# 						entidade: [endereco, potMedia, potBefore, potAfter],
# 						entidade: [endereco, potMedia, potBefore, potAfter]}
				
# 				}

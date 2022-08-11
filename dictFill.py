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

celula = ''
entidade = ''
endereco = ''
potMedia = 0
potBefore = 0
potAfter = 0

class WriteDict:

	def __init__(self):
		pass

	def fillDict (self, dadosDict, celula, entidade, endereco, potMedia, potBefore, potAfter):
		if not dadosDict.get(celula):
			dadosDict[celula] = {entidade: [endereco, potMedia, potBefore, potAfter]}
			return 'Dados adicionados!'
		else:
			# OBS: Note que se a entidade já existir, o valor dela será substituído por [endereco, potMedia, potBefore, potAfter]
			dadosDict[celula][entidade] = [endereco, potMedia, potBefore, potAfter]
			return 'Dados adicionados!'


teste = WriteDict()

dadosGerais = teste.fillDict()

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

DadosGerais = {celula: {entidade: [endereco, potMedia, potBefore, potAfter],
						entidade: [endereco, potMedia, potBefore, potAfter],
						entidade: [endereco, potMedia, potBefore, potAfter]},
				
				celula: {entidade: [endereco, potMedia, potBefore, potAfter],
						entidade: [endereco, potMedia, potBefore, potAfter],
						entidade: [endereco, potMedia, potBefore, potAfter]},
				
				celula: {entidade: [endereco, potMedia, potBefore, potAfter],
						entidade: [endereco, potMedia, potBefore, potAfter],
						entidade: [endereco, potMedia, potBefore, potAfter]}
				
				}

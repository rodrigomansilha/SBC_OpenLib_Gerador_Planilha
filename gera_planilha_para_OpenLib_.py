#!/usr/bin/python3
# -*- coding: iso-8859-15 -*-

__author__ = 'Rodrigo Mansilha'
__email__ = 'mansilha@unipampa.edu.br'
__version__ = '{0}.{0}.{1}'
__credits__ = ['Comissão Local de Pesquisa (CLP) do Campus Alegrete da Univ. Federal do Pampa - Unipampa']


# Bibliotecas gerais
import sys, os	# system
import argparse # parse arguments
import glob		# Unix style pathname pattern expansion
import logging  # organizar saídas em formato de log
from pathlib import Path

# Outras bibliotecas
from openpyxl import load_workbook, Workbook #ler e escrever arquivos em formato do Excel
import bibtexparser #ler e escrever arquivos em formato bibtex   #pip3 install bibtexparser
from bibtexparser.bparser import BibTexParser
from bibtexparser.customization import convert_to_unicode

# ARQUIVOS DE SAÍDA
PLANILHA_ARTIGOS_PADRAO = "Artigos.xlsx"
PLANILHA_AUTORES_PADRAO = "Autores.xlsx"
PLANILHA_REFERENCIAS_PADRAO = "Referencias.xlsx"
PLANILHA_SECOES_PADRAO = "Secoes.xlsx" # Não utilizada

# DIRETÓRIO DE ENTRADA
DIRETORIO_ENTRADA_PADRAO = "./papers"

# CONFIGURAÇÕES PADRÃO
LANGUAGE_PADRAO = "pt"
SECAO_ABREV_PADRAO = "ART"
AUTHOR_COUNTRY_PADRAO = "Brasil"
PLANILHA_PADRAO = "Planilha1"

# CAMPOS DAS TABELAS
CAMPOS_ARTIGOS = [
		"seq",
		"language",
		"sectionAbbrev",
		"title",
		"titleEn",
		"abstract",
		"abstractEn",
		"keywords",
		"keywordsEn",
		"pages",
		"fileLabel",
		"fileLink"
]

CAMPOS_AUTORES = [
	"article",
	"authorFirstname",
	"authorMiddlename",
	"authorLastname",
	"authorAffiliation",
	"authorAffiliationEn",
	"authorCountry",
	"authorEmail",
	"orcid",
	"authorBio",
	"authorBioEn"
]

CAMPOS_SECOES = [
	"sectionTitle",
	"sectionTitleEn",
	"sectionAbbrev"
]


class Secao(object):

	def __init__(self, sectionAbbrev_=None):
		self.sectionTitle = ""
		self.sectionTitleEn = ""
		self.sectionAbbrev = ""
		if sectionAbbrev_ is not None:
			self.sectionAbbrev = sectionAbbrev_


class Referencia(object):

	def __init__(self, artigo_seq_, referencia_seq_):

		self.article = artigo_seq_
		self.referencia_seq = referencia_seq_
		self.reference = ""
		self.DOI = ""


class Autor(object):

	id_contador = 0

	def __init__(self, artigo_seq_, autor_str_ = None, author_country_ = None):
		Autor.id_contador += 1
		self.article = artigo_seq_
		self.authorFirstname = ""
		self.authorMiddlename = ""
		self.authorLastname = ""
		self.authorAffiliation = ""
		self.authorAffiliationEn = ""
		self.authorCountry = ""
		self.authorEmail = ""
		self.orcid = ""
		self.authorBio = ""
		self.authorBioEn = ""

		if not autor_str_ is None:
			self.authorFirstname = autor_str_.split(" ")[0]
			self.authorLastname = " ".join(autor_str_.split(" ")[1:])

		if author_country_ is None:
			authorCountry = AUTHOR_COUNTRY_PADRAO

	def __str__(self):
		msg = "\n"
		msg += "\t\tAutor\n"
		msg += "\t\t-----\n"
		msg += "\t\t\tarticle: %d\n" % self.article
		msg += "\t\t\tauthorFirstname:%s\n" % self.authorFirstname
		msg += "\t\t\tauthorLastname:%s\n" % self.authorLastname
		msg += '\n'
		return msg


class Artigo(object):

	seq_contador = 0

	def __init__(self, language_, sectionAbbrev_, bib_database_= None):
		Artigo.seq_contador += 1
		self.seq = Artigo.seq_contador
		self.language = language_
		self.sectionAbbrev = sectionAbbrev_
		self.title = ""
		self.titleEn = ""
		self.abstract = ""
		self.abstractEn = ""
		self.keywords = ""
		self.keywordsEn = ""
		self.pages = ""
		self.fileLabel = ""
		self.fileLink = ""
		self.autores = []

		if not bib_database_ is None:
			for attribute in self.__dict__.keys():
				if attribute in bib_database_.entries[0]:
					value = bib_database_.entries[0][attribute]
					self.__setattr__(attribute, value)
					logging.debug("Attribute: %s [OK] Value: %s" % (attribute, value))
				else:
					logging.debug("Attribute: %s [MISS]" % attribute)

			self.fileLink = bib_database_.entries[0]["url"]

			conta_autor = 0
			autores_lista = bib_database_.entries[0]["author"].split("and")
			for autor_str in autores_lista:
				conta_autor += 1
				autor_str = autor_str.strip()
				logging.info("\t\t(%d/%d) %s" % (conta_autor, len(autores_lista), autor_str))
				autor = Autor(self.seq, autor_str)
				logging.debug(autor)
				self.autores.append(autor)

	def __str__(self):
		msg = "\n"
		msg += "\tArtigo\n"
		msg += "\t-------\n"
		msg += "\t\tseq: %d\n" % self.seq
		msg += "\t\tlanguage:%s\n" % self.language
		msg += "\t\tsectionAbbrev:%s\n" % self.sectionAbbrev
		msg += "\t\ttitle:%s\n" % self.title
		msg += "\t\ttitleEn:%s\n" % self.titleEn
		msg += "\t\tkeywords:%s\n" % self.keywords
		msg += "\t\tkeywordsEn:%s\n" % self.keywordsEn
		msg += "\t\tfileLink:%s\n" % self.fileLink
		#msg += "\tautores:%s\n" % self.autores
		for autor in self.autores:
			msg += autor.__str__()
		msg += '\n'
		return msg


def exporta_artigos_xlsx(planilha_, artigos_, acrescentar_=True):
	'''
	Preenche planilha de artigos com dados

	:param planilha_: openpyxl.sheet
	:param dados_detalhe_grupos_: dicionário com valores
	'''

	linha = gera_cabecalho(planilha_, CAMPOS_ARTIGOS, 1)
	coluna = 1

	for artigo in artigos_:
		coluna = 1
		for campo in CAMPOS_ARTIGOS:
			valor = artigo.__getattribute__(campo)
			planilha_.cell(row=linha, column=coluna).value = valor
			logging.debug("artigo: %d campo: %s valor: %s" % (artigo.seq, campo, valor))
			coluna += 1
		linha += 1

def exporta_autores_xlsx(planilha_, artigos_, acrescentar_=True):
	'''
	Preenche planilha de autores com dados

	:param planilha_: openpyxl.sheet
	:param dados_detalhe_grupos_: dicionário com valores
	'''

	linha = gera_cabecalho(planilha_, CAMPOS_AUTORES, 1)
	coluna = 1

	for artigo in artigos_:
		for autor in artigo.autores:
			coluna = 1

			for campo in CAMPOS_AUTORES:
				valor = autor.__getattribute__(campo)
				planilha_.cell(row=linha, column=coluna).value = valor
				logging.debug("artigo: %d campo: %s valor: %s" % (artigo.seq, campo, valor))
				coluna += 1
			linha += 1

def exporta_secoes_xlsx(planilha_, secao_, acrescentar_=True):
	'''
	Exporta dados de seção para planilha
	:param planilha_:
	:param secao_:
	:param acrescentar_:
	:return:
	'''

	linha = gera_cabecalho(planilha_, CAMPOS_SECOES, 3)
	coluna = 1

	for campo in CAMPOS_SECOES:
		valor = secao_.__getattribute__(campo)
		planilha_.cell(row=linha, column=coluna).value = valor
		logging.debug("campo: %s valor: %s" % (campo, valor))
		coluna += 1


def gera_cabecalho(planilha_, campos_, coluna_indice_):
	linha = 1
	coluna = 1

	if planilha_.cell(row=linha, column=coluna).value is None:
		for campo in campos_:
			planilha_.cell(row=linha, column=coluna).value = campo
			coluna += 1
		linha += 1

	coluna = coluna_indice_
	while not planilha_.cell(row=linha, column=coluna).value is None:
		linha += 1

	return linha


def gera_workbook_planilha(nome_arquivo_, acrescentar_=True):

	logging.debug("def_gera_workbook_planilha  nome_arquivo_:%s acrescentar_:%s"%(nome_arquivo_, acrescentar_))
	if Path(nome_arquivo_).is_file() and not acrescentar_:
		logging.debug("removendo arquivo...")
		os.remove(nome_arquivo_)
		logging.debug("pronto.")

	if not Path(nome_arquivo_).is_file():
		logging.debug("arquivo não existe! criando novo arquivo...")
		workbook = Workbook()
		workbook.remove(workbook.active)
		workbook.create_sheet(PLANILHA_PADRAO)
		workbook.save(nome_arquivo_)
		workbook.close()

	workbook = load_workbook(nome_arquivo_)
	planilha = workbook[PLANILHA_PADRAO]

	return workbook, planilha


def main():
	'''
	Programa principal
	'''

	# Configura argumentos
	parser = argparse.ArgumentParser(description='Gera arquivos de produção bibliográfica de eventos da SBC.')
	parser.add_argument("--dir", "-d", help="diretório com arquivos de entrada.", default=DIRETORIO_ENTRADA_PADRAO)
	parser.add_argument("--secao", "-s", help="abreviatura da seção a ser processada (coluna sectionAbbrev do arquivo Secoes.xlsx).", default=SECAO_ABREV_PADRAO)
	parser.add_argument('--acrescentar', dest='acrescentar',  help="acrescentar aos arquivos pré-existentes.", action='store_true', default=True)
	parser.add_argument('--nao-acrescentar', dest='acrescentar', help="sobreescrever arquivos pré-existentes.", action='store_false')

	parser.add_argument("--artigos", "-a", help="planilha de saída com artigos.", default=PLANILHA_ARTIGOS_PADRAO)
	parser.add_argument("--autores", "-u", help="planilha de saída com autores.", default=PLANILHA_AUTORES_PADRAO)
	parser.add_argument("--referencias", "-r", help="planilha de saída com referências.", default=PLANILHA_REFERENCIAS_PADRAO)
	parser.add_argument("--secoes", "-e", help="planilha de saída com seções.", default=PLANILHA_SECOES_PADRAO)

	help_log = "nível de log (INFO=%d DEBUG=%d)"%(logging.INFO, logging.DEBUG)
	parser.add_argument("--log", "-l", help=help_log, default=logging.INFO, type=int)
	parser.print_help()

	# lê argumentos da linha de comando
	args = parser.parse_args()

	# configura log
	#logging.basicConfig(level=args.log, format='%(asctime)s - %(message)s')
	logging.basicConfig(level=args.log, format='%(message)s')

	# mostra parâmetros de entrada
	logging.info("")
	logging.info("PARÂMETROS DE ENTRADA")
	logging.info("---------------------")

	logging.info("\tdir        : %s" % args.dir)
	logging.info("\tseção      : %s" % args.secao)
	logging.info("\tartigos    : %s" % args.artigos)
	logging.info("\tautores    : %s" % args.autores)
	logging.info("\treferencias: %s" % args.referencias)
	logging.info("\tsecoes     : %s" % args.secoes)

	logging.info("\tlog        : %d" % args.log)
	logging.info("\tacrescentar: %s" % args.acrescentar)

	# mostra parâmetros calculados
	logging.info("")
	logging.info("PARÂMETROS CALCULADOS")
	logging.info("---------------------")

	# arquivos_bib = [arquivo for arquivo in Path(args.dir).rglob('[!~]*.bib')]
	# logging.info("\tbibs   : %s" % arquivos_bib)

	arquivos_pdfs = [arquivo for arquivo in Path(args.dir).rglob('[!~]*.pdf')]
	logging.info("\tpdfs   : %s" % arquivos_pdfs)

	# inicializa variáveris
	artigos = []
	conta_arquivo = 0

	# processa dados
	logging.info("")
	logging.info("LEITURA DE DADOS")
	logging.info("----------------")

	try:
		for posix_path_pdf in arquivos_pdfs:
			conta_arquivo += 1

			nome_arquivo_pdf = str(posix_path_pdf)
			nome_arquivo_bib = nome_arquivo_pdf.replace(".pdf", ".bib")
			if not Path(nome_arquivo_bib).is_file:
				logging.exception("\t(%d/%d) PDF:%s [OK] BIB:%s [NOT FOUND] " % (conta_arquivo, len(arquivos_pdfs), nome_arquivo_pdf, nome_arquivo_bib))

			else:
				logging.info("\t(%d/%d) PDF:%s [OK] BIB:%s [OK]"%(conta_arquivo, len(arquivos_pdfs), nome_arquivo_pdf, nome_arquivo_bib))

			try:
				parser = BibTexParser()
				parser.customization = convert_to_unicode
				bibtex_file = open(nome_arquivo_bib)
				bib_database = bibtexparser.load(bibtex_file, parser=parser)
				artigo = Artigo(LANGUAGE_PADRAO, args.secao, bib_database)
				logging.info(artigo)
				artigos.append(artigo)

			except Exception as e:
				print(e)
				artigo = Artigo(LANGUAGE_PADRAO, args.secao)
				artigos.append(artigo)

	except Exception as e:
		print(e)
		#sys.exit(-1)


	logging.info("")
	logging.info("RESULTADOS")
	logging.info("----------")

	logging.info("")
	logging.info("\tSEÇÃO")
	logging.info("\t-----")
	logging.info("\t\tprocessando...")
	workbook, planilha = gera_workbook_planilha(args.secoes, args.acrescentar)
	secao = Secao(args.secao)
	exporta_secoes_xlsx(planilha, secao)
	logging.info("\t\tgravando...")
	workbook.save(args.secoes)
	logging.info("\t\tpronto.")

	logging.info("")
	logging.info("\tARTIGOS")
	logging.info("\t----------")
	logging.info("\t\tprocessando...")
	workbook, planilha = gera_workbook_planilha(args.artigos, args.acrescentar)
	exporta_artigos_xlsx(planilha, artigos)
	logging.info("\t\tgravando...")
	workbook.save(args.artigos)
	logging.info("\t\tpronto.")

	logging.info("")
	logging.info("\tAUTORES")
	logging.info("\t----------")
	logging.info("\t\tprocessando...")
	workbook, planilha = gera_workbook_planilha(args.autores, args.acrescentar)
	exporta_autores_xlsx(planilha, artigos)
	logging.info("\t\tgravando...")
	workbook.save(args.autores)
	logging.info("\t\tpronto.")

	logging.info("\nfim.\n")

if __name__ == "__main__":
	main()

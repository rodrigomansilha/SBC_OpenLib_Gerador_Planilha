# SBC_OpenLib_Gerador_Planilha
Uma ferramenta para apoio de geração de dados para o sistema OpenLib da SBC

#Exemplo para gerar arquivos a partir de diretório padrão (papers/)

python3 processa_dados_SBC.py 

#Exemplo para gerar arquivos com seções organizadas em diretórios

for d in "ERRC_completos" "ERRC_resumos" "WRSEG_completos" "WRSEG_resumos"; do python3 processa_dados_SBC.py -d $d -s $d; done; 


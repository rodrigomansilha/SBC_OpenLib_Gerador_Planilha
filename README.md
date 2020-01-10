# SBC_OpenLib_Gerador_Planilha
Uma ferramenta para gerar arquivos para o sistema OpenLib da SBC.

# Verificar comandos disponíveis

python3 gera_planilha_para_OpenLib.py -h


# Gerar arquivos a partir de diretório padrão (papers/)

python3 gera_planilha_para_OpenLib.py 

# Gerar arquivos com seções organizadas em diretórios

for d in "ERRC_completos" "ERRC_resumos" "WRSEG_completos" "WRSEG_resumos"; do python3 gera_planilha_para_OpenLib.py -d $d -s $d; done; 


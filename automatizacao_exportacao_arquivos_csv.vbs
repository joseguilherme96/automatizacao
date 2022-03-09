' Automatização, para exportar dados em arquivo csv do sistema


set shl = CreateObject("Wscript.Shell")
shl.Run "J:\ATUAL\APLICATIVOS_XE7\S2R.exe"
wscript.sleep(15000)
i=0
Do Until i=1

REM ***************FATURAMENTO POR PRODUTO *************
shl.SendKeys"%"
wscript.sleep(500)
shl.SendKeys"r"
wscript.sleep(500)
shl.SendKeys"{DOWN 5}" 
wscript.sleep(500)
shl.SendKeys"{RIGHT}" 
wscript.sleep(500)
shl.SendKeys"{DOWN 10}" 
wscript.sleep(500)
shl.sendkeys"{Enter}"
wscript.sleep(5000)
shl.sendkeys"{TAB 6}"
wscript.sleep(500)
shl.sendkeys"01022022"  'Data de inicio
wscript.sleep(500)
shl.sendkeys"{TAB}"  'Data de inicio
wscript.sleep(500)
shl.sendkeys"28022022"  'Data de fim
wscript.sleep(500)
shl.sendkeys"{TAB 3}"
wscript.sleep(500)
shl.SendKeys"C:\Users\joseg\Desktop\" 'Caminho que será salvo o arquivo
wscript.sleep(500)
shl.SendKeys"%G"
wscript.sleep(5000)
shl.SendKeys"{Enter}"
wscript.sleep(10000)
shl.SendKeys"%{TAB}"
wscript.sleep(5000)
shl.SendKeys"{Enter}"
wscript.sleep(5000)
shl.SendKeys"{ESC}"
wscript.sleep(5000)
REM ************** FIM DA EXPORTAÇÃO FATURAMENTO **************



REM ********** SUB ITENS NOTA FISCAL ********************
shl.SendKeys"%"
wscript.sleep(500)
shl.SendKeys"n"
wscript.sleep(500)
shl.SendKeys"{DOWN 5}"
wscript.sleep(500)
shl.SendKeys"{ENTER}"
wscript.sleep(500)
shl.SendKeys"{F9}"
wscript.sleep(5000)
shl.SendKeys"{F2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(500)
shl.SendKeys"{DOWN}"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.sendkeys"%m"
wscript.sleep(500)
shl.SendKeys"{DEL}"
wscript.sleep(500)
shl.sendkeys"{DOWN 2}"
wscript.sleep(500)
shl.SendKeys"+{END}"
wscript.sleep(500)
shl.SendKeys"{DEL}"
wscript.sleep(500)
shl.sendkeys"{DOWN 24}"
wscript.sleep(500)
shl.SendKeys"01022022" 'Data de cadastramento do doc
wscript.sleep(500)
shl.SendKeys"{TAB}"
wscript.sleep(500)
shl.SendKeys"e"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.SendKeys"28022022"  'Digita Ultima data de cadastro do doc
wscript.sleep(500)
shl.SendKeys"^{ENTER}"
wscript.sleep(20000)

' Exporta csv

shl.SendKeys"{F7}"
wscript.sleep(500)
shl.SendKeys"{DOWN 2}"
wscript.sleep(500)
shl.SendKeys"{ENTER 2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(5000)
shl.SendKeys"{UP 10}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{TAB 4}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ESC 2}"
wscript.sleep(5000)

REM **************** FIM DA EXPORTAÇÃO NOTA FISCAIS *********************

REM ******************** NECESSIDADE ************************

shl.sendkeys"%"
wscript.sleep(500)
shl.sendkeys"d"
wscript.sleep(500)
shl.sendkeys"{DOWN 5}"
wscript.sleep(500)
shl.sendkeys"{ENTER}"
wscript.sleep(5000)
shl.sendkeys"{F2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(500)
shl.SendKeys"{DOWN}"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.sendkeys"%m"
wscript.sleep(500)
shl.SendKeys"{DOWN}" 
wscript.sleep(500)
shl.SendKeys"PROR" 
wscript.sleep(500)
shl.SendKeys"{DOWN 22}" 
wscript.sleep(500)
shl.SendKeys"01022022"  'Digita data de cadastro
wscript.sleep(500)
shl.SendKeys"{TAB}"
wscript.sleep(500)
shl.SendKeys"e"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.SendKeys"28022022"  'Digita data de cadastro
wscript.sleep(500)
shl.SendKeys"^{ENTER}"
wscript.sleep(5000)

' Exporta csv

shl.SendKeys"{F7}"
wscript.sleep(500)
shl.SendKeys"{DOWN 2}"
wscript.sleep(500)
shl.SendKeys"{ENTER 2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(5000)
shl.SendKeys"{UP 10}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{TAB 4}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ESC}"
wscript.sleep(5000)
REM ************************FIM DE EXPORTAÇÃO NECESSIDADE************

REM *******************PEDIDO DE COMPRA***********************

shl.sendkeys"%"
wscript.sleep(500)
shl.sendkeys"c"
wscript.sleep(500)
shl.sendkeys"{DOWN 2}"
wscript.sleep(500)
shl.sendkeys"{ENTER}"
wscript.sleep(5000)
shl.sendkeys"{F2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(500)
shl.SendKeys"{DOWN}"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.sendkeys"%m"
wscript.sleep(500)
shl.SendKeys"{DOWN 10}"
wscript.sleep(500)
shl.SendKeys"+{TAB}"
wscript.sleep(500)
shl.SendKeys"c"
wscript.sleep(500)
shl.SendKeys"{TAB}"
wscript.sleep(500)
shl.SendKeys"pires" 'Digita apelido fornecedor
wscript.sleep(500)
shl.SendKeys"{TAB}"
wscript.sleep(500)
shl.SendKeys"o"
wscript.sleep(500)
shl.SendKeys"{TAB}"
wscript.sleep(500)
shl.SendKeys"c"
wscript.sleep(500)
shl.SendKeys"{TAB}"
wscript.sleep(500)
shl.SendKeys"regenfer"  'Digita apelido fornecedor
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.SendKeys"{DOWN 14}" 
wscript.sleep(500)
shl.SendKeys"01022022"  'Digita data de cadastro
wscript.sleep(500)
shl.SendKeys"{TAB}"
wscript.sleep(500)
shl.SendKeys"e"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.SendKeys"28022022"  'Digita final data de cadastro
wscript.sleep(500)
shl.SendKeys"^{ENTER}"
wscript.sleep(5000)

' Exporta csv

shl.SendKeys"{F7}"
wscript.sleep(500)
shl.SendKeys"{DOWN 2}"
wscript.sleep(500)
shl.SendKeys"{ENTER 2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(5000)
shl.SendKeys"{UP 10}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{TAB 4}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ESC}"
wscript.sleep(5000)


REM ************* FIM DA EXPORTAÇÃO DO PEDIDO DE COMPRA ********************


REM *************** PEDIDO DE VENDA *****************
shl.sendkeys"%"
wscript.sleep(500)
shl.sendkeys"v"
wscript.sleep(500)
shl.sendkeys"{DOWN 2}"
wscript.sleep(500)
shl.sendkeys"{ENTER}"
wscript.sleep(5000)
shl.sendkeys"{F2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(500)
shl.SendKeys"{DOWN}"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.sendkeys"%m"
wscript.sleep(500)
shl.SendKeys"{DOWN 28}" 
wscript.sleep(500)
shl.SendKeys"01022022"  'Digita data de cadastro
wscript.sleep(500)
shl.SendKeys"{TAB}"
wscript.sleep(500)
shl.SendKeys"e"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.SendKeys"28022022"  'Digita data final de cadastro
wscript.sleep(500)
shl.SendKeys"^{ENTER}"
wscript.sleep(5000)

' Exporta csv

shl.SendKeys"{F7}"
wscript.sleep(500)
shl.SendKeys"{DOWN 2}"
wscript.sleep(500)
shl.SendKeys"{ENTER 2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(5000)
shl.SendKeys"{UP 10}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{TAB 4}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ESC}"
wscript.sleep(5000)


REM ***************** FIM DA EXPORTAÇÃO PEDIDO DE VENDA ******************

REM ********** ORDENS ********************
shl.SendKeys"%"
wscript.sleep(500)
shl.SendKeys"{RIGHT 5}"
wscript.sleep(500)
shl.SendKeys"{DOWN 7}"
wscript.sleep(500)
shl.SendKeys"{ENTER}"
wscript.sleep(500)
shl.SendKeys"{F2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(500)
shl.SendKeys"{DOWN}"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.sendkeys"%m"
wscript.sleep(500)
shl.SendKeys"{DOWN 12}"
wscript.sleep(500)
shl.SendKeys"01022022"  'Data de Inicio
wscript.sleep(500)
shl.SendKeys"{TAB}"
wscript.sleep(500)
shl.SendKeys"e"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.SendKeys"28022022"  'Digita data final de cadastro
wscript.sleep(500)
shl.SendKeys"^{ENTER}"
wscript.sleep(500)

' Exporta csv

shl.SendKeys"{F7}"
wscript.sleep(500)
shl.SendKeys"{DOWN 2}"
wscript.sleep(500)
shl.SendKeys"{ENTER 2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(5000)
shl.SendKeys"{UP 10}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{TAB 4}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ESC}"
wscript.sleep(5000)

REM ********** NOTAS FISCAIS ********************
shl.SendKeys"%"
wscript.sleep(500)
shl.SendKeys"{RIGHT 9}"
wscript.sleep(500)
shl.SendKeys"{DOWN 6}"
wscript.sleep(500)
shl.SendKeys"{ENTER}"
wscript.sleep(500)
shl.SendKeys"{F2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(500)
shl.SendKeys"{DOWN}"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.sendkeys"%m"
wscript.sleep(500)
shl.SendKeys"{DOWN 6}"
wscript.sleep(500)
shl.SendKeys"01022022" 'Data de cadastramento do doc
wscript.sleep(500)
shl.SendKeys"{TAB}"
wscript.sleep(500)
shl.SendKeys"e"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.SendKeys"28022022"  'Digita Ultima data de cadastro do doc
wscript.sleep(500)
shl.SendKeys"^{ENTER}"
wscript.sleep(500)

' Exporta csv

shl.SendKeys"{F7}"
wscript.sleep(500)
shl.SendKeys"{DOWN 2}"
wscript.sleep(500)
shl.SendKeys"{ENTER 2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(5000)
shl.SendKeys"{UP 10}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{TAB 4}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ESC}"
wscript.sleep(5000)

REM **************** FIM DA EXPORTAÇÃO NOTA FISCAIS *********************


REM ********** SOLICITAÇÃO DE COMPRA ********************
shl.SendKeys"%"
wscript.sleep(500)
shl.SendKeys"{RIGHT 6}"
wscript.sleep(500)
shl.SendKeys"{ENTER 2}"
wscript.sleep(5000)
shl.SendKeys"{F2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(500)
shl.SendKeys"{DOWN}"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.sendkeys"%m"
wscript.sleep(500)
shl.SendKeys"{DOWN 8}"
wscript.sleep(500)
shl.SendKeys"BFT" 'Digita insumo BFT
wscript.sleep(500)
shl.SendKeys"{DOWN 13}"
wscript.sleep(500)
shl.SendKeys"01022022" 'Digita data de cadastro
wscript.sleep(500)
shl.SendKeys"{TAB}"
wscript.sleep(500)
shl.SendKeys"e"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.SendKeys"28022022"  'Digita ultima data de cadastro
wscript.sleep(500)
shl.SendKeys"^{ENTER}"
wscript.sleep(500)

' Exporta csv

shl.SendKeys"{F7}"
wscript.sleep(500)
shl.SendKeys"{DOWN 2}"
wscript.sleep(500)
shl.SendKeys"{ENTER 2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(5000)
shl.SendKeys"{UP 10}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{TAB 4}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ESC}"
wscript.sleep(5000)

REM **************** FIM DO COMANDO *********************

REM ***************** COTAÇÕES DE COMPRA ******************
shl.SendKeys"%"
wscript.sleep(500)
shl.SendKeys"c"
wscript.sleep(500)
shl.SendKeys"{DOWN}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{F2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(500)
shl.SendKeys"{DOWN}"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.sendkeys"%m"
wscript.sleep(500)
shl.SendKeys"{DOWN 18}"
wscript.sleep(500)
shl.SendKeys"01022022" 'Digita ultima data de cadastro
wscript.sleep(5000)
shl.SendKeys"{TAB}"
wscript.sleep(500)
shl.SendKeys"e"
wscript.sleep(500)
shl.SendKeys"{TAB 2}"
wscript.sleep(500)
shl.SendKeys"28022022"  'Digita ultima data de cadastro
wscript.sleep(500)
shl.SendKeys"^{ENTER}"
wscript.sleep(500)

'Exporta csv

shl.SendKeys"{F7}"
wscript.sleep(500)
shl.SendKeys"{DOWN 2}"
wscript.sleep(500)
shl.SendKeys"{ENTER 2}"
wscript.sleep(5000)
shl.SendKeys"+{TAB 3}"
wscript.sleep(5000)
shl.SendKeys"{UP 10}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{TAB 4}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ENTER}"
wscript.sleep(5000)
shl.SendKeys"{ESC}"
wscript.sleep(5000)

REM ************** FIM EXPORTAÇÃO COTAÇAÕ DE COMPRA '


REM ********** Comando para exportar arquivo csv Dsempenho fabricação divisão 2000/3500 ********************

shl.SendKeys"%"
wscript.sleep(500)
shl.SendKeys"{RIGHT 15}"
wscript.sleep(500)
shl.SendKeys"{ENTER}" 'Lista opções Ralatório
wscript.sleep(500)
shl.SendKeys"{DOWN 1}" 
wscript.sleep(500)
shl.SendKeys"{RIGHT}" 
wscript.sleep(500)
shl.SendKeys"{DOWN 18}" 
wscript.sleep(500)
shl.sendkeys"{Enter}"
wscript.sleep(5000)
shl.sendkeys"{TAB 2}"
wscript.sleep(500)
shl.sendkeys" "
wscript.sleep(500)
shl.sendkeys"{TAB 6}"
wscript.sleep(500)
shl.sendkeys"2000"
wscript.sleep(500)
wscript.sleep(500)
shl.sendkeys"{TAB 4}"
wscript.sleep(500)
shl.sendkeys"01022022"  'Data de inicio
wscript.sleep(500)
shl.sendkeys"28022022"  'Data de fim
wscript.sleep(500)
shl.sendkeys"{TAB 11}"
shl.SendKeys"C:\Users\joseg\Desktop" 'Caminho que será salvo o arquivo
wscript.sleep(500)
shl.sendkeys"{TAB 1}"
shl.SendKeys"Desempenho_Fabricacao_Div_2000.csv" 'Divisão
wscript.sleep(500)
shl.SendKeys"%G"
wscript.sleep(5000)
shl.SendKeys"%{TAB}"
wscript.sleep(5000)
shl.SendKeys"{Enter}"
wscript.sleep(5000)
shl.SendKeys"{Backspace 40}"
wscript.sleep(5000)
shl.SendKeys"Desempenho_Fabricacao_Div_3000.csv" 'Divisão
wscript.sleep(500)
shl.sendkeys"+{TAB 16}"
wscript.sleep(500)
shl.sendkeys"3500"
wscript.sleep(500)
shl.SendKeys"%G"
wscript.sleep(5000)
shl.SendKeys"{Enter}"
wscript.sleep(5000)
shl.SendKeys"%{TAB}"
wscript.sleep(5000)
shl.SendKeys"{Enter}"
wscript.sleep(5000)
shl.SendKeys"{ESC}"
wscript.sleep(5000)

REM **************** FIM DO COMANDO *********************

MsgBox("Exportação foi realizada com sucesso !")
shl.SendKeys"%{TAB}"

i=i+1
loop
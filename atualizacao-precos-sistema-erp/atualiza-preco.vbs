'Atualiza os preços dos produtos no sistema ERP 

set cmd = CreateObject("Wscript.Shell")

wscript.sleep(5000)

Dim arr(50,2) 'Tamanho da Matriz Linha x Coluna


'Coluna 0 = Produto 
'Coluna 2 = Preço de Venda

'Dados da planilha transformado em arrays conforme imagem
'Produtos
arr(9,0)="B50.020.020" ' Linha 9 Coluna 0
arr(10,0)="B50.020.026" ' Linha 10 Coluna 0...
arr(11,0)="B50.020.034"
arr(12,0)="B50.020.046"
arr(13,0)="B50.020.060"
arr(14,0)="B50.020.070"
arr(15,0)="B50.025.016"
arr(16,0)="B50.025.020"
arr(17,0)="B50.025.026"
arr(18,0)="B50.025.034"
arr(19,0)="B50.025.046"
arr(20,0)="B50.025.060"
arr(21,0)="B50.025.070"
arr(22,0)="B50.025.085"
arr(23,0)="B50.025.090"
arr(24,0)="B50.032.016"
arr(25,0)="B50.032.020"
arr(26,0)="B50.032.026"
arr(27,0)="B50.032.034"
arr(28,0)="B50.032.046"
arr(29,0)="B50.032.060"
arr(30,0)="B50.032.070"
arr(31,0)="B50.032.085"
arr(32,0)="B50.032.090"
arr(33,0)="B50.032.115"
arr(34,0)="B50.040.034"
arr(35,0)="B50.040.046"
arr(36,0)="B50.040.060"
arr(37,0)="B50.040.070"
arr(38,0)="B50.040.085"
arr(39,0)="B50.040.090"
arr(40,0)="B50.040.115"
arr(41,0)="B50.040.140"
arr(42,0)="B50.050.060"
arr(43,0)="B50.050.070"
arr(44,0)="B50.050.085"
arr(45,0)="B50.050.090"
arr(46,0)="B50.050.115"
arr(47,0)="B50.050.140"
arr(48,0)="B60.016.013"
arr(49,0)="B60.016.016"
arr(50,0)="B60.016.020"

'Preço de venda
arr(9,2)="97,55683875"
arr(10,2)="99,44458125"
arr(11,2)="101,751195"
arr(12,2)="104,61976875"
arr(13,2)="107,37297375"
arr(14,2)="110,2276875"
arr(15,2)="105,0312375"
arr(16,2)="106,67918625"
arr(17,2)="108,90429375"
arr(18,2)="111,66498"
arr(19,2)="115,22293125"
arr(20,2)="118,67165625"
arr(21,2)="122,1603075"
arr(22,2)="125,5570575"
arr(23,2)="128,53785"
arr(24,2)="117,43326"
arr(25,2)="119,672049"
arr(26,2)="122,479245"
arr(27,2)="125,854848"
arr(28,2)="130,41231"
arr(29,2)="135,635199"
arr(30,2)="139,544559"
arr(31,2)="144,9511665"
arr(32,2)="147,501879"
arr(33,2)="155,8648665"
arr(34,2)="136,778901"
arr(35,2)="142,311792"
arr(36,2)="148,638609"
arr(37,2)="153,363714"
arr(38,2)="160,028589"
arr(39,2)="162,955989"
arr(40,2)="189,728301"
arr(41,2)="205,7318235"
arr(42,2)="172,48551075"
arr(43,2)="179,72541825"
arr(44,2)="190,0533495"
arr(45,2)="194,15672325"
arr(46,2)="210,594867"
arr(47,2)="227,06766075"
arr(48,2)="85,4621355"
arr(49,2)="86,85873"
arr(50,2)="88,382406"


tempoEsperaServidor = 18000 ' Tempo espera necessário para processamento dos dados no servdor 18000 milisegundos

i=9 ' Linha do dado
x = 0 'Coluna 0 - Produto
y = 2 'Coluna 2  Preço de Venda

Do Until i=51 'Percorre dados na matriz encontrando produto e preço

cmd.SendKeys"{F2}"
wscript.sleep(500)
cmd.SendKeys"{DOWN}"
wscript.sleep(500)
cmd.SendKeys"{DEL}"
wscript.sleep(500)
cmd.SendKeys""&arr(i,x)'Localiza Produto
wscript.sleep(500)
cmd.SendKeys"^{ENTER}"
wscript.sleep(tempoEsperaServidor)
cmd.SendKeys"{F5}"
wscript.sleep(500)
cmd.SendKeys"{TAB 2}"
wscript.sleep(1000)
cmd.SendKeys""""&arr(i,y) 'Localiza valor do produto
wscript.sleep(500)
cmd.SendKeys"^{ENTER}"
wscript.sleep(500)
cmd.SendKeys"{ENTER}"
wscript.sleep(tempoEsperaServidor)

i=i+1

loop
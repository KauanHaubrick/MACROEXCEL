# MACROEXCEL
Script em VBA criado para utilizar um MACRO no Excel com o intuito de automatizar a busca dos pesos das ordens de venda no SAP.
Exemplo do funcionamento do SCRIPT.
Ele busca o valor da celular A4 (Numero da OV),
Com o SAP aberto na tela inicial, ele começa a operar abrindo a transação VA03,
Na VA03 ele cola o valor da celular A4 e acessa a OV,
No campo VALOR TOTAL ele copia o texto que contem o peso da OV e retorna com o valor para o EXCEL,
Ele copia o valor na celula D4 e continua o processo até encontrar a celula vazia ou chegar na celula A24 que é onde está o limite.

# automacao_desenhos_cod

Este programa tem por objetivo facilitar a liberação de desenhos novos e modificados para a pasta do servidor (Desenhos Fábrica).
O programa NÃO substitui a conferência final pelo usuário, a impressão dos documentos, envio dos e-mails e correção das operações.

Passo-a-passo:
	1 - abra o arquivo executável - automacao_desenhos_cod.exe;
	2 - tecle 'Enter' para iniciar o programa;
	3 - o programa irá ler os arquivos na pasta Novos Desenhos e criar uma planilha (FLUXO DESENHOS.xlsx) 	contendo os códigos, deixando em branco uma coluna para preencher com a operação respectiva;
	4 - Abra a planilha e preencha somente as células da coluna 'PROCESSOS'.
		a. Caso haja necessidade, abra o programa "Cadastro Do Roteiro de Fabricação" do Focco (FENG0202) para ter as operações;
	5 - Para preencher as células, há validação de dados, então preencher EXATAMENTE a operação de acordo com 	alguma pré-existente (entre as aspas duplas):

	"LASER", "LASER, DOBRA", "USINAGEM", "SERRA", "MONTAGEM", "PRE MONTAGEM",
  	 "TERCEIROS", "ALIMENTADORES", "SOLDA", "FONTES", "COMPRADO", "CENTRO",
  	  "OUTROS", "SERRA, CENTRO", "ROSQUEAMENTO", "GRAVADOR", "SERRA, TORNO",
   	 "SERRA, TORNO, CENTRO", "DOBRA", "SEM ROTEIRO"

	6 - Após preencher, salve a planilha e feche-a;
	7 - Retorne ao programa e tecle "Enter";
	8 - O programa irá realizar a movimentação dos arquivos conforme a versão e criará documentos de texto contendo os códigos para envio de e-mails e impressão, conforme a operação. Seguir o fluxo padrão normalmente;
		a. Desenhos sem roteiro ou outras operações não serão manipulados pelo programa. Será criado o arquivo de texto "VERIFICAR".
	9 - As operações realizadas serão descritas no ambiente do programa e escritas na coluna OBS para o respectivo código;
	10 - Copiar as informações da planilha para a planilha mestra;
	11 - Fechar o programa e EXCLUIR a planilha "FLUXO DESENHOS.xlsx".

Autor: @tpiccoli
Rev 0 - 02/2025

# MacroBaixarPedidoAutomatico
VBA EXCEL com SAP GUI

Essa macro foi feita para baixar pedidos de compras automaticamente no SAP GUI usando o VBA do Excel. 
A macro usa a transação ME9F do SAP GUI para fazer o download do pedido em PDF.
Como é algo que tira muito tempo baixar cada pedido e é muito repetitivo, fiz essa macro para tornar algo mais fácil e mais rápido, deixando mais tempo para fazer outras atividades.


Passo a passo de como usar:
1. Abra o SAP.
2. Digite o(s) número(s) do(s) pedido(s) abaixo do campo "Pedidos que estão pendentes"(pode digitar quantos números quiser, desde que seja uma linha diferente para cada número de pedido na coluna 20).
3. Clique no botão "Baixar pedido automático.".
4. Preencha o campo Login e Senha, com o seu login e senha do SAP ou clique em "Já estou logado", caso já esteja logado no SAP.
5. Clique em OK.
6. Caso o pedido já tem todas as aprovações para ser baixado, ele vai estar na seu "Área de trabalho", mas caso tiver alguma aprovação faltando vai aparecer uma mensagem na coluna da frente do número do pedido no 
   excel de quem está faltando a aprovação (abaixo do campo "O que falta para enviar").

Obs:
Muito provavelmente o código vai precisar de pequenos ajustes para funcionar, pois o meu SAP muito provavelmente tem pequenas diferenças do seu, já que nem todas as empresas adotam as mesmas configurações do SAP.
Caso seja necessário alguma mudança, eu deixei no código do VBA o que cada parte do código está fazendo, para ficar mais fácil de identificar possíveis erros.

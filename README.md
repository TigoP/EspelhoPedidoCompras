Relatório Espelho de Pedido de Compras
Este projeto contém o código fonte para a geração do relatório de "Espelho de Pedido de Compras" utilizando TL++ (ADVPL) no ambiente Protheus. O relatório exibe informações consolidadas do pedido, tais como o cabeçalho, os itens principais (com detalhes em duas seções: dados da tabela e informações descritivas) e os totais.

Sumário
Visão Geral

Arquivos Principais

Estrutura do Relatório

Funcionalidades

Como Usar

Personalização e Ajustes

Contribuição

Licença

Visão Geral
O relatório de espelho de pedido de compras foi desenvolvido para auxiliar na visualização detalhada dos dados de um pedido, incluindo os itens e seus detalhes, além dos totais e condições de pagamento, transporte, dentre outras informações. O código está estruturado para ser modular, facilitando manutenções e futuras alterações.

Arquivos Principais
EspelhoPedido.prw – Arquivo principal que contém a função EspelhoPedido() e a configuração inicial do objeto gráfico (oImpFRMC).

Funções de Impressão – Conjunto de funções:

fGetDados() / fGetMock() – Responsáveis por obter (ou gerar dados de teste) o conteúdo do pedido.

fImprimeItens() – Função que imprime os itens do pedido. Ela divide a impressão em duas seções:

A primeira parte exibe uma tabela com os dados principais (número do item, código, NCM, CFOP, data de entrega, quantidade, valor unitário, desconto e total).

A segunda parte imprime os detalhes descritivos de cada item (descrição, fornecedor, código externo, observação, solicitado, embalagem e prazo de entrega).

fImprimeTotais() – Função que imprime os totais do pedido (valor dos produtos, frete, seguro, impostos, descontos e valor total) organizados em colunas.

Outras funções de impressão para as seções de transporte e condições de pagamento também foram implementadas.

Estrutura do Relatório
O relatório é composto das seguintes partes:

Cabeçalho Geral
Contém informações do pedido e configuração do objeto gráfico (oImpFRMC), tais como resolução, orientação e margens.

Itens do Pedido

Tabela Principal:
Imprime os dados básicos do item em formato tabular com colunas fixas:

Colunas: #, Item, NCM, CFOP, Entrega, Quantidade, Vlr Un, Vlr Desc, Vlr Total

Detalhes do Item:
Imprime abaixo da linha principal os detalhes do item, como:

Descrição, Fornecedor, Código Externo, Observação, Solicitado, Embalagem e Entrega.

Totais do Pedido
Os totais são impressos em formato de colunas, mostrando os valores de:

Valor Total dos Produtos, Frete, Seguro, Outras Despesas, Desconto, e, para cada imposto (ex.: IPI, PIS, COFINS, ICMS), a Base de Cálculo e o Valor do Imposto.

Seções Adicionais
Outras seções, como transporte e condições de pagamento, podem ser incluídas conforme necessário, utilizando a mesma abordagem modular e iterativa.

Funcionalidades
Modularização:
Cada função tem uma responsabilidade única, de forma que a impressão de cada parte do relatório é facilmente manipulável e ajustável.

Iteração Dinâmica:
Tanto os itens quanto os totais são impressos utilizando arrays e loops, garantindo que o sistema lide com qualquer quantidade de itens sem alterações manuais no código.

Configuração Gráfica:
Utiliza o objeto FWMSPrinter para configurar a impressão (resolução, orientacão, tamanho do papel, margens, etc.) e gerar o PDF.

Suporte a Dados Mock:
Em ambiente de teste, as funções fGetDados() e fGetMock() permitem a geração de dados fictícios para visualizar e validar o relatório.

Como Usar
Configuração do Ambiente:
Certifique-se de que o ambiente Protheus esteja devidamente configurado com TL++ (ADVPL) e que as bibliotecas necessárias (ex.: fwprintsetup.ch, rptdef.ch, tlpp-core.th) estejam disponíveis.

Execução:
Chame a função EspelhoPedido() para gerar o relatório de espelho de pedido de compras. Essa função se encarrega de:

Inicializar o objeto gráfico (oImpFRMC)

Obter os dados (reais ou mock)

Chamar as funções de impressão de cada parte do relatório (cabeçalho, itens, totais, transporte, etc.)

Finalizar a página e exibir uma pré-visualização, se configurado.

Customização:
Para alterar o layout, basta ajustar os offsets (por exemplo, em aCols ou aColOffsets) ou os espaçamentos verticais (nRow_ += ...) conforme necessário.

Personalização e Ajustes
Ajuste de Layout:
Os arrays que definem os offsets das colunas (como aCols e aColOffsets) podem ser modificados para alterar a distribuição horizontal das colunas.

Espaçamento Vertical:
Variáveis como nRow_ são incrementadas com valores fixos (por exemplo, 15, 20, 40) para controlar o espaçamento entre linhas. Esses valores podem ser ajustados conforme seu design.

Alteração de Fontes:
As fontes utilizadas (ex.: oFont12, oFont10) podem ser alteradas para ajustar o tamanho e o estilo conforme o padrão da empresa.

Contribuição
Contribuições são bem-vindas!
Se você encontrar melhorias, correções ou quiser adicionar funcionalidades, sinta-se à vontade para abrir issues ou submeter pull requests.

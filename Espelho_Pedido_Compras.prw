#include "totvs.ch"
#include "fwprintsetup.ch"
#include "rptdef.ch"
//#include "tlpp-core.th"

#DEFINE ALIGN_H_LEFT 0
#DEFINE ALIGN_H_RIGHT 1
#DEFINE ALIGN_H_CENTER 2
#DEFINE ALIGN_V_CENTER 0

//namespace Compras.Relatorio

Static lMock__ := .F. as logical

/*/
{Protheus.doc} EspelhoPedido
(Imprime Espelho Pedido de Compra)
@return logical, .T. se impressão bem sucedida
/*/

User function EspelhoPedido() // Compras.Relatorio.U_EspelhoPedido

    Local lDisabeSetup := .T.  as logical
    Local aDados       := {}   as array

    Private nRow_      := 0    as numeric
    Private nCol_      := 0    as numeric
    Private nColTot_   := 0    as numeric
    Private nRowTot_   := 0    as numeric
    Private nCntPag_   := 0    as numeric
    Private nSizePage_ := 0    as numeric
    Private nMargRel_  := 60   as numeric
    Private cEspLinha_ := "-8" as character

    Private cBitmap            as character
    Private oImpFRMC_          as object
    Private oFont10_           as object
    Private oFont12_           as object
    Private oFont20_           as object
    Private oFont12n_          as object
    Private oFont14n_          as object
    Private oFont20n_          as object
    Private oBrush_            as object

    lPreview           := .T.

    oImpFRMC_          := FWMSPrinter():New('FRMC', IMP_PDF, , , lDisabeSetup)

    oImpFRMC_:SetResolution(72)
    oImpFRMC_:SetPortrait()
    oImpFRMC_:SetPaperSize(DMPAPER_A4)
    oImpFRMC_:SetMargin(10,10,10,10)
    oImpFRMC_:nDevice  := IMP_PDF
    oImpFRMC_:cPathPDF := "C:\temp\"
    oImpFRMC_:lInJob   := .F.

    cBitmap_   := "/system/lgrlt_esp_pc.bmp"
    nSizePage_ := oImpFRMC_:nPageWidth
    nMargRel_  := 60
    nRow_      := nMargRel_
    nCol_      := nMargRel_

    nColTot_   := nSizePage_ - (nMargRel_ + nMargRel_)
    nRowTot_   := oImpFRMC_:nPageHeight - (nMargRel_ + nMargRel_)

    oFont8_    := TFont() :New("Arial", 9         , 8 , .T., .F., 5, .T., 5, .T., .F.)
    oFont10_   := TFont() :New("Arial", 9         , 10, .T., .F., 5, .T., 5, .T., .F.)
    oFont12_   := TFont() :New("Arial", 9         , 12, .T., .F., 5, .T., 5, .T., .F.)
    oFont20_   := TFont() :New("Arial", 9         , 20, .T., .F., 5, .T., 5, .T., .F.)
    oFont12n_  := TFont() :New("Arial", 9         , 12, .T., .T., 5, .T., 5, .T., .F.)
    oFont14n_  := TFont() :New("Arial", 9         , 14, .T., .T., 5, .T., 5, .T., .F.)
    oFont20n_  := TFont() :New("Arial", 9         , 20, .T., .T., 5, .T., 5, .T., .F.)

    oBrush_    := TBrush():New(       , CLR_YELLOW)

    if fGetDados(@aDados)

        fPrint(aDados)

        oImpFRMC_:EndPage()

        If lPreview
            oImpFRMC_:Preview()
        EndIf

    endif

    FreeObj(oImpFRMC_) // Libera a memória
    oImpFRMC_ := Nil // Apaga o objeto

return(.T.)

/*/
    {Protheus.doc} fGetDados
    (Relatório de pedido de compras )
/*/
Static Function fGetDados(aDados as array)

    Local jCabecalho   := JsonObject():New() as json
    Local aItens       := {}                 as array
    Local jTotais      := JsonObject():New() as json
    Local jCobranca    := JsonObject():New() as json
    Local jTransporte  := JsonObject():New() as json
    Local jObservacao  := JsonObject():New() as json
    Local jNotasRodape := JsonObject():New() as json
    Local jRodape      := JsonObject():New() as json
    Local jPergunte                          as json

    Local lResult as logical

    aDados := { jCabecalho, aItens, jTotais, jCobranca, jTransporte, jObservacao, jNotasRodape, jRodape }

    if lMock__

        fGetMock(@aDados)

        lResult := .T.

    else

        if( lResult := fPergunte( @jPergunte ) )

            fGetFromDataBase(@aDados, jPergunte)

        endif

    endif

Return lResult

/*/{Protheus.doc} fPergunte
Função para perguntar parâmetros ao usuário.
@type function
@param jRetParam, json, Retorna os parâmetros conforme definição do usuário.
@return logical, Retorna .T. (true) para indicar que o usuário confirmou a operação e .F. (false) caso contrário..
/*/
Static Function fPergunte( jRetParam as json )

    Local aParam     := {}                 as array
    Local lResult    := .F.                as logical
    Local cInitPar1                        as character
    Local cInitPar2                        as character
    Local cInitPar3                        as character
    Local cInitPar4                        as character
    Local aComboMail :={" ", "Sim", "Não"} as array
    Local xComboMail                       as variant

    If FunName() == "MATA121"
        cInitPar1 := SC7->C7_NUM
        cInitPar2 := SC7->C7_NUM
        cInitPar3 := SC7->C7_EMISSAO
        cInitPar4 := SC7->C7_EMISSAO

    else
        cInitPar1 := PadL( '' ,6, ' ' )
        cInitPar2 := PadR( '' ,6, ' ' )
        cInitPar3 := cToD("//")
        cInitPar4 := cToD("//")

    endif

    // [1]-Tipo 1 -> MsGet()
    // [2]-Descricao
    // [3]-String contendo o inicializador do campo
    // [4]-String contendo a Picture do campo
    // [5]-String contendo a validacao
    // [6]-Consulta F3
    // [7]-String contendo a validacao When
    // [8]-Tamanho do MsGet
    // [9]-Flag .T./.F. Parametro Obrigatorio ?
    aadd(aParam, {1, "Pedido de"       , cInitPar1, "@!"      , ".T.", "SC7", ".T.", 55, .T.})
    aadd(aParam, {1, "Pedido Até"      , cInitPar2, "@!"      , ".T.", "SC7", ".T.", 55, .T.})
    aadd(aParam, {1, "Dt. Emissão De"  , cInitPar3,           , ".T.",      , ".T.", 55, .T.})
    aadd(aParam, {1, "Dr. Emissão Até" , cInitPar4,           , ".T.",      , ".T.", 55, .T.})

    // [2]-Tipo 2 -> Combo
    // [2]-Descricao
    // [3]-Numerico contendo a opcao inicial do combo
    // [4]-Array contendo as opcoes do Combo
    // [5]-Tamanho do Combo
    // [6]-Validacao
    // [7]-Flag .T./.F. Parametro Obrigatorio ?
    aadd(aParam, {2, "Envia por E-Mail", 2        , aComboMail, 45   , ".T.", .T.})

    while .T.

        // 1 - <aParametros> - Vetor com as configurações
        // 2 - <cTitle>      - Título da janela
        // 3 - <aRet>        - Vetor passador por referencia que contém o retorno dos parâmetros
        // 4 - <bOk>         - Code block para validar o botão Ok
        // 5 - <aButtons>    - Vetor com mais botões além dos botões de Ok e Cancel
        // 6 - <lCentered>   - Centralizar a janela
        // 7 - <nPosX>       - Se não centralizar janela coordenada X para início
        // 8 - <nPosY>       - Se não centralizar janela coordenada Y para início
        // 9 - <oDlgWizard>  - Utiliza o objeto da janela ativa
        //10 - <cLoad>       - Nome do perfil se caso for carregar
        //11 - <lCanSave>    - Salvar os dados informados nos parâmetros por perfil
        //12 - <lUserSave>   - Configuração por usuário
        aRetParam := {}
        if ParamBox( aParam, "Parâmetros", @aRetParam,/*bOk*/,/*aButtons*/,/*lCentered*/,/*nPosX*/,/*nPosY*/,;
                /*oDlgWizard*/,/*cLoad*/,/*lCanSave*/.F.,/*lUserSave*/.F.)

            xComboMail                  := aRetParam[5]

            jRetParam                   := JsonObject():New()
                jRetParam["pedidoDe"]       := aRetParam[1]
                jRetParam["pedidoAte"]      := aRetParam[2]
                jRetParam["dataEmissaoDe"]  := aRetParam[3]
                jRetParam["dataEmissaoAte"] := aRetParam[4]
                jRetParam["enviaPorEmail"]  := iif( valType(xComboMail) == 'C' , xComboMail, aComboMail[xComboMail] )

            lResult                     := .T.

            exit

        elseif FwAlertYesNo("Deseja realmente cancelar?", "Cancelar")
            exit
        endif

    enddo

Return lResult

/*/{Protheus.doc} fGetFromDataBase
Obtem dados para impressão a partir da Base de Dados.
@type function
@param aDados, array, Retorna os dados para impressão.
@param jPergunte, json, Parâmetros de usuário.
/*/
Static Function fGetFromDataBase( aDados as array, jPergunte as json )

    Local aArea           := GetArea()                   as array
    Local aAreaSE4        := SE4->(GetArea())            as array
    Local aAreaSA2        := SA2->(GetArea())            as array
    Local aAreaSC7        := SC7->(GetArea())            as array

    Local cPedidoDe       := jPergunte["pedidoDe"]       as character
    Local cPedidoAte      := jPergunte["pedidoAte"]      as character
    Local cDataEmissaoDe  := jPergunte["dataEmissaoDe"]  as character
    Local cDataEmissaoAte := jPergunte["dataEmissaoAte"] as character

    Local jCabecalho      := aDados[1]
    Local jCobranca       := aDados[4]
    Local jTransporte     := aDados[5]
    Local jObservacao     := aDados[6]
    Local jNotasRodape    := aDados[7]
    Local jRodape         := aDados[8]

    Private aItens_               := aDados[2]
    Private jTotais_              := aDados[3]
            jTotais_["valorTotalProduto"] := 0
            jTotais_["frete"]             := 0
            jTotais_["seguro"]            := 0
            jTotais_["outrasDespesas"]    := 0
            jTotais_["desconto"]          := 0
            jTotais_["ipi"]               := 0
            jTotais_["pis"]               := 0
            jTotais_["cofins"]            := 0
            jTotais_["icms"]              := 0
            jTotais_["bcIpi"]             := 0
            jTotais_["bcPis"]             := 0
            jTotais_["bcCofins"]          := 0
            jTotais_["bcIcms"]            := 0
            jTotais_["valorTotal"]        := 0

    Private cAlias_               := GetNextAlias() as character
    Private cPedido_                                as character

    if select("SC7") <= 0
        dbSelectArea("SA2")
    endif

    if select("SA2") <= 0
        dbSelectArea("SA2")
    endif

    BeginSql Alias cAlias_

        %noParser%

        SELECT
            SC7.R_E_C_N_O_ as C7_RECNO,
            SA2.R_E_C_N_O_ as A2_RECNO,
            SE4.R_E_C_N_O_ as E4_RECNO
        FROM
            %table:SC7% SC7
        JOIN
            %table:SA2% SA2 ON (
                SA2.A2_FILIAL = %xFilial:SA2%
                AND SA2.A2_COD = SC7.C7_FORNECE
                AND SA2.A2_LOJA = SC7.C7_LOJA
                AND SA2.D_E_L_E_T_ = ' ' )
        JOIN
            %table:SE4% SE4 ON (
                SE4.E4_FILIAL = %xFilial:SE4%
                AND SE4.E4_CODIGO = SC7.C7_COND
                AND SE4.D_E_L_E_T_ = ' ' )
        WHERE
            SC7.C7_FILIAL = %xFilial:SC7%
            AND SC7.C7_NUM BETWEEN %exp:cPedidoDe% AND %exp:cPedidoAte%
            AND SC7.C7_EMISSAO BETWEEN %exp:cDataEmissaoDe% AND %exp:cDataEmissaoAte%
            AND SC7.D_E_L_E_T_ = ' '
    endSql

    dbSelectArea(cAlias_)

    if !(cAlias_)->(eof())

        SC7->(dbGoTo((cAlias_)->C7_RECNO))
        SA2->(dbGoTo((cAlias_)->A2_RECNO))
        SE4->(dbGoTo((cAlias_)->E4_RECNO))

        while !(cAlias_)->(eof())

            cPedido_                 := SC7->C7_NUM

            jCabecalho["numero"]     := SC7->C7_NUM
            jCabecalho["emissao"]    := SC7->C7_EMISSAO
            jCabecalho["empresa"]    := SM0->M0_FULNAME
            jCabecalho["cnpj"]       := SM0->M0_CGC
            jCabecalho["ie"]         := SM0->M0_INSC
            jCabecalho["cep"]        := SM0->M0_CEPENT
            jCabecalho["endereco"]   := SM0->("Rua " + AllTrim(M0_ENDENT) + ", " + "Bairro " + AllTrim(M0_BAIRENT) +;
                                             " - " + AllTrim(M0_CIDENT) + " - " + AllTrim(M0_ESTENT) + " - " + AllTrim(M0_CEPENT))
            jCabecalho["telefone"]   := SM0->M0_TEL
            // Fornecedor
            jCabecalho["fornecedor"] := SA2->A2_NOME
            jCabecalho["ieForn"]     := SA2->A2_INSCR
            jCabecalho["endForn"]    := SA2->(AllTrim(A2_END) + ", " + AllTrim(A2_BAIRRO) + ", " + AllTrim(A2_EST) + ", " + AllTrim(A2_CEP))
            jCabecalho["telForn"]    := SA2->A2_TEL
            
            Private cMoeda_ as character

            Do Case
                Case SC7->C7_MOEDA == 1
                    cMoeda_ := "R$"
                Case SC7->C7_MOEDA == 2
                    cMoeda  := "U$"
                Case SC7->C7_MOEDA == 4
                    cMoeda  := "EUR"
                Otherwise
                    cMoeda  := SC7->C7_MOEDA
            EndCase

            jCabecalho["moeda"]         := cMoeda_
            jCobranca["formaPagamento"] := SC7->C7_COND
            jCobranca["condPagamento"]  := SE4->E4_DESCRI
            jObservacao["observacao"]   := SC7->(AllTrim(C7_OBS))
            jNotasRodape["notasRodape"] := "Supervisao.compras.es@emailho.com.br - Funcionario"
            jRodape["rodapeEmpresa"]    := AllTrim(SM0->M0_FULNAME) + " - " + "CNPJ: " + AllTrim(SM0->M0_CGC)
            jRodape["rodapeEndereco"]   := SM0->("Rua " + AllTrim(M0_ENDENT) + ", " + "Bairro " + AllTrim(M0_BAIRENT) + " - " +;
                                                 AllTrim(M0_CIDENT) + " - " + AllTrim(M0_ESTENT) + " - " + AllTrim(M0_CEPENT))
            jRodape["rodapeContato"]    := "(27) 999999999 cobranca@empresa.com.br"

            Private cCondFrete_ as Character
            
            Do Case
                Case SC7->C7_TPFRETE == "C"
                    cCondFrete_ := "CIF"
                Case SC7->C7_TPFRETE == "F"
                    cCondFrete_ := "FOB"
                Case SC7->C7_TPFRETE == "T"
                    cCondFrete_ := "Por conta de Terceiros"
                Case SC7->C7_TPFRETE == "R"
                    cCondFrete_ := "Por conta do Remetente"
                Case SC7->C7_TPFRETE == "D"
                    cCondFrete_ := "Por conta do Destinatário"
                Otherwise
                    cCondFrete_ := SC7->C7_TPFRETE
            EndCase

            jTransporte["condFrete"]      := cCondFrete_
            jTransporte["transportadora"] := SC7->(AllTrim(C7_TRANSP))

            fProductIterator()

        enddo

    endif

    (cAlias_)->(dbCloseArea())

    RestArea(aAreaSA2)
    RestArea(aAreaSE4)
    RestArea(aAreaSC7)
    RestArea(aArea)

Return

/*/{Protheus.doc} fProductIterator
Função aulixar recursiva para iterar os itens do pedido de compra.
/*/
Static Function fProductIterator()

    Local jItens as json

    jItens                        := JsonObject():New()
    jItens["item"]                := SC7->C7_PRODUTO
    jItens["ncm"]                 := SB1->B1_POSIPI
    jItens["cfop"]                := "???"
    jItens["dtEntrega"]           := SC7->C7_DATPRF
    jItens["quant" ]              := SC7->C7_QUANT
    jItens["vlrUnit" ]            := Round(SC7->C7_PRECO,2)
    jItens["vlrDesc" ]            := Round(SC7->C7_VLDESC,2)
    jItens["vlrTotal"]            := Round(SC7->C7_TOTAL,2)
    jItens["vlrIpi" ]             := Round(SC7->C7_VALIPI,2)
    jItens["aliquotaIpi"]         := SC7->C7_IPI
    jItens["vlrIcms" ]            := Round(SC7->C7_VALICM,2)
    jItens["aliquotaIcms"]        := SC7->C7_PICM
    jItens["descricao"]           := SC7->C7_DESCRI
    jItens["fornecedor"]          := SA2->A2_NOME
    jItens["codigoExterno"]       := "AR168047-0002"
    jItens["observacao"]          := "1 entrega do 2 pedido " + SA2->A2_NOME
    jItens["solicitacao"]         := 6.91
    jItens["embalagem"]           := 1
    jItens["przEntrega"]          := SC7->C7_DATPRF
    aadd(aItens_, jItens)

    jTotais_["valorTotalProduto"] := Round(SC7->C7_TOTAL ,2)
    jTotais_["frete"]             := Round(SC7->C7_FRETE ,2)
    jTotais_["seguro"]            := Round(SC7->C7_SEGURO ,2)
    jTotais_["outrasDespesas"]    := Round(SC7->C7_DESPESA,2)
    jTotais_["desconto"]          := Round(SC7->C7_VLDESC ,2)
    jTotais_["ipi"]               := Round(SC7->C7_VALIPI ,2)
    jTotais_["pis"]               := Round(SC7->C7_VALPIS,2)
    jTotais_["cofins"]            := Round(SC7->C7_VALCOF,2)
    jTotais_["icms"]              := Round(SC7->C7_VALICM ,2)
    jTotais_["bcIpi"]             := Round(SC7->C7_BASEIPI,2)
    jTotais_["bcPis"]             := Round(SC7->C7_BASPIS ,2)
    jTotais_["bcCofins"]          := Round(SC7->C7_BASCOF ,2)
    jTotais_["bcIcms"]            := Round(SC7->C7_BASEICM,2)
    jTotais_["valorTotal"]        := Round(SC7->C7_TOTAL,2) + (Round(SC7->C7_FRETE ,2) + Round(SC7->C7_SEGURO ,2) + Round(SC7->C7_DESPESA,2) + Round(SC7->C7_VALIPI ,2) - Round(SC7->C7_VLDESC ,2)) //TODO: PRECISA VER QUAL TABELA É A CORRETA

    (cAlias_)->(dbSkip())
        SC7->(dbGoTo((cAlias_)->C7_RECNO))
        SA2->(dbGoTo((cAlias_)->A2_RECNO))
        SE4->(dbGoTo((cAlias_)->E4_RECNO))

    if SC7->C7_NUM == cPedido_
        fProductIterator()
    endif

Return

/*/
    {Protheus.doc} fGetMock
    (Estrutura do Relatório de pedido de compras )
/*/
Static Function fGetMock(aDados as array)

    Local jCabecalho   := aDados[1]
    Local jItens
    Local aItens       := aDados[2]
    Local jTotais      := aDados[3]
    Local jCobranca    := aDados[4]
    Local jTransporte  := aDados[5]
    Local jObservacao  := aDados[6]
    Local jNotasRodape := aDados[7]
    Local jRodape      := aDados[8]

    jCabecalho["numero"]          := "005912"
    jCabecalho["emissao"]         := date()
    jCabecalho["empresa"]         := "Empresa Industria e Comercio Ltda"
    jCabecalho["cnpj"]            := "99999999999999"
    jCabecalho["ie"]              := "9999999999"
    jCabecalho["cep"]             := "29150410"
    jCabecalho["endereco"]        := "Rodovia da empresa, 398, Itaciba - ES"
    jCabecalho["telefone"]        := "2799999999"
    jCabecalho["fornecedor"]      := "Fornecedor Industria e Comercio Ltda"
    jCabecalho["ieForn"]          := "9999999999999"
    jCabecalho["endForn"]         := "Rua Gda empresa, 314, Bairro Santo Amaro - Sao Paulo - SP Cep 04755070"
    jCabecalho["telForn"]         := "119999999"
    jCabecalho["moeda"]           := "R$"

    jItens                        := JsonObject():New()
    jItens["item"]                := "MPC0112600021"
    jItens["ncm"]                 := "33029019"
    jItens["dtEntrega"]           := cToD("14/04/25")
    jItens["quant"]               := 7.00 //kg
    jItens["vlrUnit"]             := 285.15
    jItens["vlrDesc"]             := 0.00
    jItens["vlrTotal"]            := 1996.05
    jItens["vlrIpi"]              := 64.87
    jItens["aliquotaIpi"]         := 3.25
    jItens["vlrIcms"]             := 139.72
    jItens["aliquotaIcms"]        := 7.00
    jItens["descricao"]           := "Frag Sedosa s/ale ar168047-0002-bt skin - 11025347"
    jItens["fornecedor"]          := "Fornecedor"
    jItens["codigoExterno"]       := "AR168047-0002"
    jItens["observacao"]          := "1 entrega do 2 pedido"
    jItens["solicitacao"]         := 6.91 //kg
    jItens["embalagem"]           := 1 //kg
    jItens["przEntrega"]          := "Entregar até 14/04/25"
    aadd(aItens, jItens)

    jItens                        := JsonObject():New()
    jItens["item"]                := "MPC0112600021"
    jItens["ncm"]                 := "33029019"
    jItens["cfop"]                := "2101"
    jItens["dtEntrega"]           := cToD("14/04/25")
    jItens["quant"]               := 7.00 //kg
    jItens["vlrUnit"]             := 285.15
    jItens["vlrDesc"]             := 0.00
    jItens["vlrTotal"]            := 1996.05
    jItens["vlrIpi"]              := 64.87
    jItens["aliquotaIpi"]         := 3.25
    jItens["vlrIcms"]             := 139.72
    jItens["aliquotaIcms"]        := 7.00
    jItens["descricao"]           := "Frag Sedosa s/ale ar168047-0002-bt skin - 11025347"
    jItens["fornecedor"]          := "Fornecedor"
    jItens["codigoExterno"]       := "AR168047-0002"
    jItens["observacao"]          := "2 entrega do 2 pedido"
    jItens["solicitacao"]         := 5.76 //kg
    jItens["embalagem"]           := 1 //kg
    jItens["przEntrega"]          := "Entregar até 13/06/25"
    aadd(aItens, jItens)

    jTotais["valorTotalProduto"]  := 3706.95
    jTotais["frete"]              := 0.00
    jTotais["seguro"]             := 0.00
    jTotais["outrasDespesas"]     := 0.00
    jTotais["desconto"]           := 0.00
    jTotais["ipi"]                := 120.47
    jTotais["pis"]                := 56.88
    jTotais["cofins"]             := 262.01
    jTotais["icms"]               := 59.48
    jTotais["bcIpi"]              := 3706.95
    jTotais["bcPis"]              := 3706.95
    jTotais["bcCofins"]           := 3447.47
    jTotais["bcIcms"]             := 3706.95
    jTotais["valorTotal"]         := 3827.40

    jCobranca["formaPagamento"]   := "a prazo"
    jCobranca["condPagamento"]    := "28 dias"

    jTransporte["condFrete"]      := "1 - FOB"
    jTransporte["transportadora"] := 0
    jTransporte["pesoBruto"]      := 0.00
    jTransporte["pesoLiquido"]    := 0.00

    jObservacao["observacao"]   := "Lorem Ipsum é simplesmente um texto modelo da indústria tipográfica e de impressão. Lorem Ipsum tem sido o texto modelo padrão da indústria desde os anos 1500, quando um impressor."
    jNotasRodape["notasRodape"] := "Supervizsao.compras.es@empresa.com.br - Funcionario; Supervizsao.compras.es@empresa.com.br - Funcionario"

    jRodape["rodapeEmpresa"]    := "Empresa Industria e Comercio Ltda - CNPJ: 99999999999999"
    jRodape["rodapeEndereco"]   := "Rua: Rodovia da empresa, 398, Itaciba - ES - CEP: 29150410"
    jRodape["rodapeContato"]    := "(27) 999999999 cobranca@empresa.com.br"

Return

/*/
    {Protheus.doc} fPrint
    (Relatório de pedido de compras nova pagina )
    @type function
/*/
Static Function fPrint(aDados as array)

    Local jCabecalho   := aDados[1]
    Local aItens       := aDados[2]
    Local jTotais      := aDados[3]
    Local jCobranca    := aDados[4]
    Local jTransporte  := aDados[5]
    Local jObservacao  := aDados[6]
    Local jNotasRodape := aDados[7]
    Local jRodape      := aDados[8]

    Local nTotItens   as numeric
    Local nPrinted    as numeric
    Local nPending    as numeric
    Local nStart      as numeric
    Local nEnd        as numeric
    Local lFirstPage  as logical
    Local nFooterPos  as numeric
    Local nPagesItems as numeric
    Local nTotalPag   as numeric

    nFooterPos := nRowTot_ - 220

    oImpFRMC_:StartPage()

    nRow_      := nMargRel_
    nCol_      := nMargRel_
    nCol_R     := nColTot_ * 0.20
    fImprimeLogo(nRow_, nCol_, nCol_R)

    nCol_      := nCol_R
    nCol_R     := nColTot_ * 0.80
    fImprimeTituloPedido(jCabecalho, nRow_, nCol_, nCol_R)

    nTotItens  := Len(aItens)
    nPrinted   := 0
    nPending   := 0
    nStart     := 0
    nEnd       := 0
    lFirstPage := .F.

    Local nDiferenca := nTotItens - 6
    Local nDivisao   := nDiferenca / 8
    Local nInt       := Int(nDivisao)
    Local nMod       := nDiferenca % 8

    If nTotItens <= 6
       nPagesItems := 1
    Else
        If nMod > 0 //pega pagina 1 + qtd de paginas completas + 1 página extra para o que sobrar
            nPagesItems := 1 + nInt + 1
        Else
            nPagesItems := 1 + nInt
        EndIf
    EndIf

    nTotalPag := nPagesItems + 1 // + 1 página extra para os totais

    while .T.

        nCntPag_++

        lFirstPage := ( nCntPag_ == 1 )

        if( lFirstPage )
            nRow_      := fImprimeInfoCabecalho(jCabecalho, nRow_, nCol_, nCol_R)
            nStart     := 1
            nEnd       := iif( nTotItens < 6, nTotItens, 6 )
        else
            oImpFRMC_:StartPage()
            nRow_      := nMargRel_
            nCol_      := nMargRel_
            nCol_R     := nColTot_ * 0.20
            fImprimeLogo(nRow_, nCol_, nCol_R)

            nCol_      := nCol_R
            nCol_R     := nColTot_ * 0.80
            fImprimeTituloPedido(jCabecalho, nRow_, nCol_, nCol_R)

            nRow_      := nMargRel_ + 120
            nStart     := ( nEnd + 1 )
            nEnd       := iif( nPending < 8, nPending, 8 )
        endif

        nRow_      := fImprimeItens(aItens, nRow_, nCol_, nCol_R, nStart, nEnd)

        nRow_      := nFooterPos
        fImprimeRodape(jRodape, nRow_, nCol_, nColTot_, nCntPag_, nTotalPag)

        oImpFRMC_:EndPage()

        nPrinted   := ( nEnd - nStart + 1 )
        nPending   := nTotItens - nPrinted

        if nPending <= 0
            exit
        endif

    enddo

    nCntPag_++ // incrementa a contagem de páginas para a ultima página

    oImpFRMC_:StartPage() // Inicializa a última página

    nRow_  := nMargRel_
    nCol_  := nMargRel_
    nCol_R := nColTot_ * 0.20
    fImprimeLogo(nRow_, nCol_, nCol_R)

    nCol_  := nCol_R
    nCol_R := nColTot_ * 0.80
    fImprimeTituloPedido(jCabecalho, nRow_, nCol_, nCol_R)
    nRow_  := nMargRel_ + 300

    nRow_  := fImprimeTotais(jTotais, nRow_, nCol_, nCol_R)
    nRow_  := fImprimeCobranca(jCobranca, nRow_, nCol_, nCol_R)
    nRow_  := fImprimeTransporte(jTransporte, nRow_, nCol_, nCol_R)
    nRow_  := fImprimeObservacao(jObservacao, nRow_, nCol_, nCol_R)
    nRow_  := fImprimeNotasRodape(jNotasRodape, nRow_, nCol_, nCol_R)

    nRow_  := nFooterPos
    fImprimeRodape(jRodape, nRow_, nCol_, nColTot_, nCntPag_, nTotalPag)

    oImpFRMC_:EndPage() // Encerra a última página

Return

/*/
    {Protheus.doc} fImprimeLogo
    (Relatório de pedido de compras LOGO)
/*/
Static Function fImprimeLogo(nRow_ as numeric, nCol_ as numeric, nCol_R as numeric)

    Local nWidth  := 437 as numeric
	Local nHeight := 168 as numeric

    oImpFRMC_:Box(nRow_, nCol_, nRow_+180, nCol_+nCol_R, cEspLinha_)
    oImpFRMC_:SayBitmap( nRow_+7, nCol_+7, cBitmap_, nWidth, nHeight )

Return

/*/
    {Protheus.doc} fGetMock
    (Relatório de pedido de compras - Titulo da Pagina)
/*/
Static Function fImprimeTituloPedido(jCabecalho as json, nRow_ as numeric, nCol_ as numeric, nCol_R as numeric)

    oImpFRMC_:Box(nRow_, nCol_, nRow_+180, nCol_+nCol_R, cEspLinha_)
    oImpFRMC_:SayAlign(nRow_+60, nCol_, "Pedido de Compra - " + jCabecalho["numero"], oFont20n_, nCol_R, 30, ALIGN_V_CENTER, ALIGN_H_CENTER)

Return

/*/
    {Protheus.doc} fImprimeInfoCabecalho
    (Relatório de pedido de compras - Cabeçalho do pedido de compra)
/*/
Static Function fImprimeInfoCabecalho(jCabecalho as json, nRow_ as numeric, nCol_ as numeric, nCol_R as numeric)

    nRow_   := 240
    nCol_   := nMargRel_
    //oImpFRMC_:SayAlign(nRow_, nCol_, "Referencia: " + jCabecalho["referencia"], oFont12n_, nColTot_, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)
    oImpFRMC_:SayAlign(nRow_, nCol_, "Emissão: " + Transform(jCabecalho["emissao"], "99/99/99"), oFont12n_, nColTot_, 15, ALIGN_V_CENTER, ALIGN_H_RIGHT)

    nRow_   := 50
    nCol_   := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, jCabecalho["empresa"], oFont14n_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    nRow_   := 37
    nCol_   := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, "CNPJ: " + jCabecalho["cnpj"], oFont12_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    nRow_   := 37
    nCol_   := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, "IE: " + jCabecalho["ie"], oFont12_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    nRow_   := 37
    nCol_   := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, "CEP: " + jCabecalho["cep"], oFont12_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    nRow_   := 37
    nCol_   := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, "Endereço: " + jCabecalho["endereco"], oFont12_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    nRow_   := 37
    nCol_   := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, "Tel.: " + jCabecalho["telefone"], oFont12_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    // Fornecedor
    nRow_   := 37
    oImpFRMC_:Line(nRow_ + 40, nMargRel_, nRow_+40, nColTot_)

    nRow_   := 50
    nCol_   := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, "Fornecedor: " + jCabecalho["fornecedor"], oFont14n_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    nRow_   := 37
    nCol_   := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, "IE: " + jCabecalho["ieForn"], oFont12_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    nRow_   := 37
    nCol_   := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, "Endereço: " + jCabecalho["endForn"], oFont12_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    nRow_   := 37
    nCol_   := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, "Tel.: " + jCabecalho["telForn"], oFont12_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    nRow_   := 37
    nCol_   := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, "Moeda: " + jCabecalho["moeda"], oFont12_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

Return nRow_

/*/
    {Protheus.doc} fImprimeLinhaPontilhada
    (Relatório de pedido de compras - Cria uma linha pontilhada)
    @returns Linha pontilhada 
/*/
Static Function fImprimeLinhaPontilhada(nRow_ as numeric, nColIni as numeric, nColFim as numeric)
    Local nColAtual := nColIni
    Local cSegmento := "."

    While nColAtual <= nColFim
         oImpFRMC_:Say(nRow_, nColAtual, cSegmento, oFont8_)
         nColAtual       := 6
    Enddo
Return

/*/{Protheus.doc} fFormatarValores
    Formatação de valores utilizando mascara para impressão de milhares e centavos
    @return valor formatado
/*/
Static Function fFormatarValores(nValor)
    cTemp := StrTran(AllTrim(Transform(nValor, "999,999,999.99")), ",", "#")
    cTemp := StrTran(cTemp, ".", ",")
    cTemp := StrTran(cTemp, "#", ".")
Return cTemp

/*/
    {Protheus.doc} fImprimeItens
    (Relatório de pedido de compras - Array de objetos que itera sobre os valores do pedido de compra formando colunas e posicionando o svalores)
    @returns Os valores dos itens em colunas e titulos
/*/
Static Function fImprimeItens(aItens as array, nRow_ as numeric, nCol_ as numeric, nCol_R as numeric, nStart as numeric, nEnd as numeric) as numeric

    Local nItem, nIndex, aCols, nTotalCols
    Local aHeader, aRow, aCamposDesc
    Local cQuant, cVlrUnit, cVlrDesc, cVlrTotal, cTemp

    nRow_ := 20
    oImpFRMC_:Line(nRow_ + 40, nMargRel_, nRow_ + 40, nColTot_) // Linha no meio

    nRow_ += 50
    aCols      := { nMargRel_,; // Divisão das colunas para cabeçalho
               nMargRel_ + ((nColTot_ - nMargRel_)/8)*1, ;
               nMargRel_ + ((nColTot_ - nMargRel_)/8)*2, ;
               nMargRel_ + ((nColTot_ - nMargRel_)/8)*3, ;
               nMargRel_ + ((nColTot_ - nMargRel_)/8)*4, ;
               nMargRel_ + ((nColTot_ - nMargRel_)/8)*5, ;
               nMargRel_ + ((nColTot_ - nMargRel_)/8)*6, ;
               nMargRel_ + ((nColTot_ - nMargRel_)/8)*7 }
    nTotalCols := Len(aCols)

    aHeader    :={"#", "Item", "NCM", "Entrega", "Quantidade", "Vlr Un", "Vlr Desc", "Vlr Total"}
    For nIndex := 1 To nTotalCols
        nWidth     := Iif(nIndex < nTotalCols, aCols[nIndex+1] - aCols[nIndex], nColTot_ - aCols[nIndex])
        oImpFRMC_:SayAlign(nRow_, aCols[nIndex], aHeader[nIndex], oFont10_, nWidth, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)
    Next

    nRow_ += 30
    oImpFRMC_:Line(nRow_ + 40, nMargRel_, nRow_ + 40, nColTot_)
    nRow_ += 40
  
    For nItem := nStart To nEnd

        // Aplica a máscara nos valores numéricos para exibir, por exemplo, "10,00" ao invés de "10"
        cQuant    := fFormatarValores(aItens[nItem]["quant"])
        cVlrUnit  := fFormatarValores(aItens[nItem]["vlrUnit"])
        cVlrDesc  := fFormatarValores(aItens[nItem]["vlrDesc"])
        cVlrTotal := fFormatarValores(aItens[nItem]["vlrTotal"])

        aRow := { AllTrim(Str(nItem)),;
                   aItens[nItem]["item"],;
                   aItens[nItem]["ncm"],;
                   dToC(aItens[nItem]["dtEntrega"]),;
                   cQuant,;
                   cVlrUnit,;
                   cVlrDesc,;
                   cVlrTotal }

        For nIndex := 1 To nTotalCols // colunas do item
            nWidth     := Iif(nIndex < nTotalCols, aCols[nIndex+1] - aCols[nIndex], 160)
            oImpFRMC_:SayAlign(nRow_, aCols[nIndex], aRow[nIndex], oFont12_, nWidth, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)
        Next

        nRow_ += 30
        aCamposDesc := { { "Descrição: ",    "descricao" },;
                         { "Fornecedor: ",   "fornecedor" },;
                         { "Código Externo: ", "codigoExterno" },;
                         { "Observação: ",     "observacao" },;
                         { "Solicitado: ",     "solicitacao" },;
                         { "Embalagem: ",      "embalagem" },;
                         { "Entrega: ",        "przEntrega" } }

        For nIndex := 1 To Len(aCamposDesc)
             nRow_      := 35
             // Se o campo for numérico, aplica a máscara; caso contrário, utiliza a conversão padrão
             If aCamposDesc[nIndex][2] == "solicitacao" .or. aCamposDesc[nIndex][2] == "embalagem"
                 cTemp      := fFormatarValores(aItens[nItem][ aCamposDesc[nIndex][2] ])
             Else
                 cTemp      := cValToChar(aItens[nItem][ aCamposDesc[nIndex][2] ])
             EndIf
             oImpFRMC_:SayAlign(nRow_, nMargRel_, aCamposDesc[nIndex][1] + cTemp, oFont12_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)
        Next

        fImprimeLinhaPontilhada(nRow_ + 40, nMargRel_, nColTot_)
        nRow_ := 50

    Next

Return nRow_

/*/
    {Protheus.doc} fImprimeTotais
    (Relatório de pedido de compras - Array de objetos que itera sobre os valores do pedido de compra)
    @returns Os valores do pedido de compra
/*/
Static Function fImprimeTotais(jTotais as json, nRow_ as numeric, nCol_ as numeric, nCol_R as numeric) as numeric

    nRow_                := 40
    oImpFRMC_:SayAlign(nRow_, nMargRel_, "Totais:", oFont14n_, nColTot_, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    nRow_                := 80
    Local nEspLinhaTotal := 40 // Espaçamento entre linhas
    Local aCols          :={nMargRel_, (nColTot_ / 10) * 2, (nColTot_ / 10) * 6, (nColTot_ / 10) * 7, (nColTot_ / 10) * 9}

    Local cValorTotalProduto, cFrete, cSeguro, cOutrasDespesas, cDesconto
    Local cIpi, cPis, cCofins, cIcms, cBcIpi, cBcPis, cBcCofins, cBcIcms, cValorTotal

    cValorTotalProduto := "R$: " + fFormatarValores(jTotais["valorTotalProduto"])
    cFrete             := fFormatarValores(jTotais["frete"])
    cSeguro            := fFormatarValores(jTotais["seguro"])
    cOutrasDespesas    := fFormatarValores(jTotais["outrasDespesas"])
    cDesconto          := fFormatarValores(jTotais["desconto"])
    cIpi               := fFormatarValores(jTotais["ipi"])
    cPis               := fFormatarValores(jTotais["pis"])
    cCofins            := fFormatarValores(jTotais["cofins"])
    cIcms              := fFormatarValores(jTotais["icms"])
    cBcIpi             := fFormatarValores(jTotais["bcIpi"])
    cBcPis             := fFormatarValores(jTotais["bcPis"])
    cBcCofins          := fFormatarValores(jTotais["bcCofins"])
    cBcIcms            := fFormatarValores(jTotais["bcIcms"])
    cValorTotal        := "R$: " + fFormatarValores(jTotais["valorTotal"])

    Local aRows := { ;
         {"Valor Total do Produtos", cValorTotalProduto, "Imposto", "Base calculo", "Valor"}, ;
         {"Frete:"                 , cFrete            , "IPI"    , cBcIpi        , cIpi}, ;
         {"Seguro:"                , cSeguro           , "PIS"    , cBcPis        , cPis}, ;
         {"Outras Despesas:"       , cOutrasDespesas   , "COFINS" , cBcCofins     , cCofins}, ;
         {"Desconto:"              , cDesconto         , "ICMS"   , cBcIcms       , cIcms}, ;
         {"Valor Total"            , cValorTotal       , ""       , ""            , ""} ;
    }

    Local nTotalCols := Len(aCols)
    Local nRowIndex, nColIndex, nWidth, cValue

    For nRowIndex := 1 To Len(aRows) // Percorre linha e coluna para impressão dos totais
        For nColIndex := 1 To nTotalCols
            cValue        := aRows[nRowIndex][nColIndex]
            nWidth        := Iif(nColIndex < nTotalCols, aCols[nColIndex+1] - aCols[nColIndex], 150)
            oImpFRMC_:SayAlign(nRow_, aCols[nColIndex], cValue, oFont12_, nWidth, nEspLinhaTotal, ALIGN_V_CENTER, ALIGN_H_LEFT)
        Next
        nRow_         := nEspLinhaTotal
    Next

Return nRow_

/*/
    {Protheus.doc} fImprimeCobranca
    (Relatório de pedido de compras - Condição de pagamento e forma de pagamento)
    @returns As condições de pagamento e forma de pagamento do pedido de compra
/*/
Static Function fImprimeCobranca(jCobranca as json, nRow_ as numeric, nCol_ as numeric, nCol_R as numeric)

    nRow_       := 40
    oImpFRMC_:Line(nRow_ + 40, nMargRel_, nRow_+40, nColTot_)

    nRow_       := 50
    oImpFRMC_:SayAlign(nRow_, nMargRel_, "Cobrança:", oFont14n_, nColTot_, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    nRow_       := 60
    Local nCol1 := nMargRel_
    Local nCol2 := nCol1 + (nColTot_ / 2)

    oImpFRMC_:SayAlign(nRow_, nCol1, "Forma de Pagamento: " + jCobranca["formaPagamento"], oFont12_, nColTot_ / 2, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)
    oImpFRMC_:SayAlign(nRow_, nCol2, "Condição de Pagamento: " + jCobranca["condPagamento"], oFont12_, nColTot_ / 2, 15, ALIGN_V_CENTER, ALIGN_H_CENTER)

Return nRow_

/*/
    {Protheus.doc} fImprimeTransporte
    (Relatório de pedido de compras - Dados de transporte do pedido de compra)
    @returns Os dados de transporte do pedido de compra
/*/
Static Function fImprimeTransporte(jTransporte as json, nRow_ as numeric, nCol_ as numeric, nCol_R as numeric)

    nRow_       := 40
    oImpFRMC_:Line(nRow_ + 40, nMargRel_, nRow_+40, nColTot_)

    nRow_       := 50
    oImpFRMC_:SayAlign(nRow_, nMargRel_, "Transporte:", oFont14n_, nColTot_, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    nRow_       := 60
    nCol_       := nMargRel_
    Local nCol1 := nCol_
    oImpFRMC_:SayAlign(nRow_, nCol1, "Transportador: " + cValToChar(jTransporte["transportadora"]), oFont12_, 600, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

    nRow_       := 35
    //Local nCol2 := (nColTot_ / 4) * 2
    //Local nCol3 := (nColTot_ / 4) * 3

    oImpFRMC_:SayAlign(nRow_, nCol_, "Condição de Frete: " + jTransporte["condFrete"], oFont12_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)

Return nRow_

/*/
    {Protheus.doc} fImprimeObservacao
    (Relatório de pedido de compras - Observações do pedido de compra)
    @returns As observações do pedido de compra
/*/
Static Function fImprimeObservacao(jObservacao as json, nRow_ as numeric, nCol_ as numeric, nCol_R as numeric)

    nRow_           := 40
    oImpFRMC_:Line(nRow_ + 40, nMargRel_, nRow_+40, nColTot_)

    nRow_           := 80
    nCol_           := nMargRel_

    Local cTexto    := jObservacao["observacao"]
    Local nMaxChars := 160
    Local nTotal    := Len(cTexto)
    Local nLinhas   := CEILING(nTotal / nMaxChars)
    Local nI, cLinha

    oImpFRMC_:SayAlign(nRow_, nCol_, "Observações:", oFont12n_, nColTot_, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)
    nRow_           := 20

    For nI := 1 To nLinhas
         cLinha := SubStr(cTexto, ((nI - 1) * nMaxChars) + 1, nMaxChars)
         nRow_  := 40
         oImpFRMC_:SayAlign(nRow_, nCol_, cLinha, oFont12_, nColTot_, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)
    Next

Return nRow_

/*/
    {Protheus.doc} fImprimeNotasRodape
    (Relatório de pedido de compras - Notas de rodapé do pedido de compra)
    @returns As notas de rodapé do pedido de compra
/*/
Static Function fImprimeNotasRodape(jNotasRodape as json, nRow_ as numeric, nCol_ as numeric, nCol_R as numeric)

    nRow_           := 40
    oImpFRMC_:Line(nRow_ + 40, nMargRel_, nRow_+40, nColTot_)

    nRow_           := 80
    nCol_           := nMargRel_

    Local cTexto    := jNotasRodape["notasRodape"]
    Local nMaxChars := 150
    Local nTotal    := Len(cTexto)
    Local nLinhas   := CEILING(nTotal / nMaxChars)
    Local nI, cLinha

    oImpFRMC_:SayAlign(nRow_, nCol_, "Por gentileza, colocar para recebimentos de boletos, NFs XML, laudos para respectivos e-mails:", oFont12_, nCol_R, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)
    nRow_           := 20

    For nI          := 1 To nLinhas
         cLinha          := SubStr(cTexto, ((nI - 1) * nMaxChars) + 1, nMaxChars)
         nRow_           := 35
         oImpFRMC_:SayAlign(nRow_, nCol_, cLinha, oFont12_, nColTot_, 15, ALIGN_V_CENTER, ALIGN_H_LEFT)
    Next

Return nRow_

/*/{Protheus.doc} fImprimeRodape
    Rodapé fixo do relatório
    @version 1.0
/*/
Static Function fImprimeRodape(jRodape as json, nRow_ as numeric, nCol_ as numeric, nColTot_ as numeric, nCntPag_ as numeric, nTotalPag_ as numeric)

    nRow_ := 40
    oImpFRMC_:Line(nRow_ + 40, nMargRel_, nRow_+40, nColTot_)

    nRow_ := 40
    nCol_ := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, jRodape["rodapeEmpresa"], oFont12n_, nColTot_, 15, ALIGN_V_CENTER, ALIGN_H_CENTER)

    nRow_ := 35
    nCol_ := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, jRodape["rodapeEndereco"], oFont12_, nColTot_, 15, ALIGN_V_CENTER, ALIGN_H_CENTER)

    nRow_ := 35
    nCol_ := nMargRel_
    oImpFRMC_:SayAlign(nRow_, nCol_, jRodape["rodapeContato"], oFont12_, nColTot_, 15, ALIGN_V_CENTER, ALIGN_H_CENTER)

    //contagem de páginas
    oImpFRMC_:SayAlign(nRow_ + 10, nMargRel_, "Página " + AllTrim(Str(nCntPag_)) + " de " + AllTrim(Str(nTotalPag_)), oFont10_, nColTot_, 15, ALIGN_V_CENTER, ALIGN_H_RIGHT)

Return nRow_

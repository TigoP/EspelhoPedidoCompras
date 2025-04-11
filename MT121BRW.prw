#include "totvs.ch"
#include "fwmvcdef.ch"

/*/                                                                                                            {Protheus.doc} MT121BRW
PE para adicionar botões no browser principal da rotina MATA120 - Pedido de Compra.
@type function
@version 12.1.2310
@author Tiago
@since 01/04/2025
/*/
User Function MT121BRW()

    aadd(aRotina, {'# Imprimir Pedido de Compra', "Compras.Relatorio.U_EspelhoPedido()", 0, MODEL_OPERATION_VIEW})

Return

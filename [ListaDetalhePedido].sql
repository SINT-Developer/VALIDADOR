USE [SRPP]
GO

/****** Object:  StoredProcedure [dbo].[ListaDetalhePedido]    Script Date: 05/28/2026 11:10:24 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE procedure [dbo].[ListaDetalhePedido] 
	@nropedido int,
   @ordem int = 0
as
begin
	declare @errmsg1 varchar(1024)=''
	declare @errmsg2 varchar(1024)=''
	declare @errmsg varchar(1024)=''

	set nocount on

	if not exists (select * from empresa where codempresa=1)
	begin
		set @errmsg1='Não foi possível detalhar Pedidos porque Empresa não foi configurada.'
		goto error
	end

   if (@ordem=0)
   begin
      select     
	      I.CodProduto,
	      Coalesce(P.CodAuxiliarProduto,'') As CodAuxiliarProduto, 
	      P.Produto, 
	      I.QtdeVendida, 
	      I.PrecoTabelaH, 
	      Replace(Convert(VarChar(8),Convert(Decimal(8,2),I.PrecoTabelaH)),'.',',') As PrecoTabelaH$, 

	      Replace(Convert(VarChar(11),Convert(Decimal(11,2),(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	      Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	      (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	      -
	      (Round((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	      Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	      (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	      * (Case When Coalesce(I.PrecoPromocionalH,'N')='N' Then Coalesce(C.DescontoCascataH,0.00)/100 Else 0.00 End), 2)))),'.',',') As SubTotal2$,

	      Replace(Convert(VarChar(11),Convert(Decimal(11,2),((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	      Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	      (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	      -
	      (Round((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	      Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	      (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	      * (Case When Coalesce(I.PrecoPromocionalH,'N')='N' Then Coalesce(C.DescontoCascataH,0.00)/100 Else 0.00 End), 2)))

	      +
	      (Round(((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	      Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	      (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	      -
	      (Round((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	      Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	      (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	      * (Case When Coalesce(I.PrecoPromocionalH,'N')='N' Then Coalesce(C.DescontoCascataH,0.00)/100 Else 0.00 End), 2))) * Coalesce(I.AliquotaIPIH,0.00) / 100,2)))),'.',',') As SubTotal3$

	      from
	      cabecalhopedido c, itempedido i, produto p
	      where
	      c.nropedido= @nropedido and
	      c.nropedido = i.nropedido and
	      p.codproduto = i.codproduto
	      order by
         i.dthrinclusao desc
   end
   else
   begin
      if (@ordem=1)
      begin
         select     
	         I.CodProduto,
	         Coalesce(P.CodAuxiliarProduto,'') As CodAuxiliarProduto, 
	         P.Produto, 
	         I.QtdeVendida, 
	         I.PrecoTabelaH, 
	         Replace(Convert(VarChar(8),Convert(Decimal(8,2),I.PrecoTabelaH)),'.',',') As PrecoTabelaH$, 

	         Replace(Convert(VarChar(11),Convert(Decimal(11,2),(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	         Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	         (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	         -
	         (Round((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	         Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	         (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	         * (Case When Coalesce(I.PrecoPromocionalH,'N')='N' Then Coalesce(C.DescontoCascataH,0.00)/100 Else 0.00 End), 2)))),'.',',') As SubTotal2$,

	         Replace(Convert(VarChar(11),Convert(Decimal(11,2),((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	         Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	         (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	         -
	         (Round((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	         Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	         (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	         * (Case When Coalesce(I.PrecoPromocionalH,'N')='N' Then Coalesce(C.DescontoCascataH,0.00)/100 Else 0.00 End), 2)))
	         +
	         (Round(((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	         Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	         (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	         -
	         (Round((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	         Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	         (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	         * (Case When Coalesce(I.PrecoPromocionalH,'N')='N' Then Coalesce(C.DescontoCascataH,0.00)/100 Else 0.00 End), 2))) * Coalesce(I.AliquotaIPIH,0.00) / 100,2)))),'.',',') As SubTotal3$
	         FROM
	         CabecalhoPedido C, ItemPedido I, Produto P
	         Where
	         C.NroPedido= @NroPedido And
	         C.NroPedido = I.NroPedido And
	         P.CodProduto = I.CodProduto
	         Order By
            P.Produto Asc
      end
      else
      begin
         select     
	         I.CodProduto,
	         Coalesce(P.CodAuxiliarProduto,'') As CodAuxiliarProduto, 
	         P.Produto, 
	         I.QtdeVendida, 
	         I.PrecoTabelaH, 
	         Replace(Convert(VarChar(8),Convert(Decimal(8,2),I.PrecoTabelaH)),'.',',') As PrecoTabelaH$, 

	         Replace(Convert(VarChar(11),Convert(Decimal(11,2),(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	         Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	         (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	         -
	         (Round((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	         Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	         (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	         * (Case When Coalesce(I.PrecoPromocionalH,'N')='N' Then Coalesce(C.DescontoCascataH,0.00)/100 Else 0.00 End), 2)))),'.',',') As SubTotal2$,

	         Replace(Convert(VarChar(11),Convert(Decimal(11,2),((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	         Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	         (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	         -
	         (Round((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	         Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	         (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	         * (Case When Coalesce(I.PrecoPromocionalH,'N')='N' Then Coalesce(C.DescontoCascataH,0.00)/100 Else 0.00 End), 2)))
	         +
	         (Round(((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	         Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	         (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	         -
	         (Round((Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2) -
	         Round(Round(I.QtdeVendida*Coalesce(I.PrecoTabelaH,0.00),2)*Coalesce(I.DescontoIndividual,0.00)/100,2) -
	         (Case When Coalesce(I.MultiploGradeH,0) > 0 And Coalesce(I.DescontoGradeH,0) > 0 Then  Round(Round((I.QtdeVendida - (I.QtdeVendida % I.MultiploGradeH)) * Coalesce(I.PrecoTabelaH,0.00),2) * Coalesce(I.DescontoGradeH,0.00) /100,2)  Else 0.00 End))
	         * (Case When Coalesce(I.PrecoPromocionalH,'N')='N' Then Coalesce(C.DescontoCascataH,0.00)/100 Else 0.00 End), 2))) * Coalesce(I.AliquotaIPIH,0.00) / 100,2)))),'.',',') As SubTotal3$
	         FROM
	         CabecalhoPedido C, ItemPedido I, Produto P
	         Where
	         C.NroPedido= @NroPedido And
	         C.NroPedido = I.NroPedido And
	         P.CodProduto = I.CodProduto
	         Order By
            I.CodProduto Asc
      end
   end

	return 0

error:
	set @errmsg=dbo.setmsg2(@errmsg1, @errmsg2)
	raiserror (@errmsg, 16 ,1)
	return -1
END

GO



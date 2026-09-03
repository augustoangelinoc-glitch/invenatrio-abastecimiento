
ABASTECIMIENTO – VERSIÓN FINAL

MAPEO REAL

STOCK
Codigo
Descripción
U.Medida
Sistema -> Stock actual
Familia
Costo Cierre Mes -> Costo unitario

SALIDAS
Material -> material/código
F.Almac -> fecha de salida
Unidades -> cantidad salida
(En la estructura indicada, F.Almac es H y Unidades es AX.)

ORDENES DE COMPRA
Material -> material/código
F. Docum. -> fecha de documento
Fecha Guia -> recepción
Unidades -> cantidad comprada
Estado -> filtro COMPRADO

REGLAS IMPORTANTES
1. El código se conserva como texto.
2. Se utiliza una llave secundaria numérica únicamente para cruzar casos donde Excel
   haya eliminado ceros iniciales en una fuente. El código mostrado sigue siendo
   exactamente el del Stock.
3. Consumo mensual = suma de salidas del mes.
4. Consumo total = suma de todos los meses analizados.
5. Meses con consumo = meses cuyo consumo > 0.
6. Promedio mensual = consumo total / meses con consumo.
7. Consumo diario = promedio mensual / 30.
8. Días sin movimiento = fecha de análisis - última salida.
9. Cobertura = stock actual / consumo diario.
10. Lead Time = Fecha Guia - F. Docum., solo para OC COMPRADO y valores 0–365 días.
11. No se inventa Lead Time cuando no existe una conexión válida.
12. Última compra = última OC COMPRADO.
13. Compra 1/2/3 meses = promedio mensual x meses.
14. Duración después de comprar = (stock + compra) / consumo diario.
15. Cantidad a abastecer = stock objetivo - stock actual, nunca negativa.
16. Stock objetivo = promedio mensual x horizonte seleccionado + stock de seguridad.
17. Punto de pedido = demanda durante Lead Time + stock de seguridad.
18. Si no hay consumo, no se divide entre cero.
19. Un material con una sola salida se clasifica como Ocasional.
20. Se mantienen todas las columnas definidas en la tabla del proyecto.
21. Se genera Excel descargable.
22. No se escribe en st.session_state después de instanciar widgets.

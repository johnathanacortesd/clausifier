---
name: comparar-portales
description: Compara cómo cubren a una marca o cliente cinco portales de noticias colombianos (temas, tono para la marca y diferencias entre portales) con las notas más recientes de la web. Úsala cuando pidan buscar noticias de un cliente en portales o medios colombianos, comparar portales, ver el tono por medio o armar un comparativo de cobertura.
---

# Comparar portales

Lanza cinco ayudantes `analista-portal` en paralelo, uno por portal, y une lo que devuelven en
`entregas/comparativo.md` con el tema y el tono para la marca en cada portal y en qué se
diferencia la información entre ellos.

## Procedimiento

1. Confirma el CLIENTE (marca). Si no lo dieron, pídelo y para. Toma también sus ALIAS
   (siglas, nombres alternos) y VOCEROS (personas que hablan por la marca) si los dieron: una
   nota que nombra un alias o un vocero cuenta como nota del cliente.
2. Usa estos cinco portales, salvo que el usuario nombre otros: El Tiempo (eltiempo.com),
   El Espectador (elespectador.com), Semana (semana.com), Infobae Colombia
   (infobae.com/colombia) y Caracol Radio (caracol.com.co).
3. En UN solo mensaje, lanza cinco veces la herramienta `Agent` con el ayudante
   `analista-portal`, una por portal, para que corran en paralelo. A cada uno pásale: CLIENTE,
   ALIAS, VOCEROS y su dominio. No busques tú en la web mientras ellos trabajan.
4. Lee las cinco respuestas. Si dos portales cubren el mismo hecho, unifica el TEMA para que
   sea exactamente el mismo en los dos. No cambies el tono que trajo el ayudante salvo que
   contradiga las reglas de TONO; si lo cambias, dilo.
5. Si un ayudante devolvió "sin notas recientes", registra el portal así en la tabla. No
   lances otro ayudante para el mismo portal.
6. Escribe `entregas/comparativo.md` con `Write`, con esta forma:
   - `# Comparativo de portales: <cliente>` y la fecha de hoy.
   - Tabla: Portal | Fecha | Titular | TEMA | TONO | Enlace (una fila por nota).
   - Tabla resumen: Portal | Nº notas | Temas principales | Tono dominante.
   - `## Cómo se diferencia la información`: 3 a 6 viñetas concretas (qué hecho destaca cada
     portal, qué fuentes o voceros cita, cuál es más crítico o más favorable a la marca).
   - Una frase final de conclusión.

## Reglas

- No inventes notas, fechas ni enlaces. Cada fila lleva la URL real que trajo un ayudante. Si
  no la tienes, la fila no va.
- "Lo más reciente" significa las notas con fecha más nueva que encontraron. Escribe la fecha
  como aparece; si no aparece, pon "sin fecha".
- El tono se juzga para el cliente, no por lo grave o alegre que sea la noticia. Lo que dicen
  sus voceros, o lo que se dice de ellos en su rol, cuenta para el tono del cliente.
- Solo escribes dentro de `entregas/`.
- Cierra con UNA frase de conclusión.

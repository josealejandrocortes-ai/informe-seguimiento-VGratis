# Informe Seguimiento Semanal — diligenciamiento automático

Aplicación que completa el informe semanal de Word (cadena presupuestal) a partir del Excel de soporte,
conservando el formato del Word.

## Instalación y uso
```bash
pip install -r requirements.txt
streamlit run app.py
```
1. Sube la **plantilla Word** (con marcadores) y el **Excel de soporte** (.xlsm/.xlsx).
2. Revisa la fecha de corte y la lista de valores que se insertarán.
3. (Opcional) Elige el modelo en la barra lateral, genera los análisis `<…>` y edítalos.
4. Genera y descarga el informe.

Sin interfaz: `python -m informe_semanal.cli PLANTILLA.docx SOPORTE.xlsm -o INFORME.docx [--sin-ia]`
(para los análisis define la variable de la clave del proveedor, p. ej. `GEMINI_API_KEY`, y usa `--proveedor`).

> La aplicación lee los **valores guardados** en el Excel. Antes de subirlo: actualiza las tablas dinámicas de
> `TD Maestro`, recalcula y guarda.

## Sintaxis de marcadores en el Word
| Marcador | Resultado |
|---|---|
| `{'TD Maestro'!B2}` | Valor de la celda **tal como lo muestra Excel** (usa el formato de número de la celda). Vale en párrafos y en celdas de tablas. |
| `{'Hoja'!B2\|M}` | Con modificador. `M` pesos→`$8.767M` · `MM` ya en millones→`$8.767M` · `mill` pesos→`8.767` · `pct` `0,6%` · `pct0` · `pct2` · `pp` puntos porcentuales (`0,2`) · `abs` · `dir` (aumento/disminución) · `up` · `cap` · `txt` · `num0/1/2`. Se pueden encadenar: `{'Hoja'!M13\|abs,pp}` |
| `{@fecha_corte\|larga}` | Fecha de corte: `larga` (18 de septiembre de 2026), `dm`, `ab` (18 Sep), `mes`, `iso`. `{@fecha_anterior\|…}` = 7 días antes. `{@anio}` |
| `{@mov.obligaciones.dep_art}` | Mayor variación **porcentual** semanal por momento (`compromisos`, `obligaciones`, `pagos`), calculada comparando `Ejec Presupuestal Vigente` con `Ejec Presupuestal Sem Ant`. Campos: `dep`, `dep_art`, `valor`, `pct`, `rubro`, `rubro_valor`, `rubro_desc`, `tipo` |
| `{@buscar('Hoja'!Q18; 'Flujos caja contratos'!B:C)}` | Búsqueda tipo BUSCARV (3.er argumento opcional: columna a devolver) |
| `<instrucción …>` | Análisis redactado por IA. La instrucción puede citar rangos (`'Hoja'!A1:C9`) y llevar `{…}` dentro; se envía ese contexto y un resumen calculado. Se puede insertar en medio de una frase. |

Los nombres de hoja toleran mayúsculas y espacios finales (`'Rezago presupuestal '`).

## Gráficos
Los gráficos del Word son gráficos nativos vinculados al Excel: la aplicación reescribe sus datos desde las mismas
referencias que ya traen (`'Ejecución presupuestal'!$J$19:$M$19`, etc.), por lo que **formato, tamaño y posición
no cambian**. No hace falta marcador.

## Mantenimiento mensual
Los marcadores apuntan a celdas fijas. Cuando cambia el mes, revisa las referencias que dependen de la fila del
mes en `Legalizaciones` (`B38`, `C38`, `C42`, `C43`, `B59` en la plantilla actual).

## Modelos de IA para los análisis `<…>`
Se elige en la barra lateral. Con Streamlit Cloud, guarda la clave en **Secrets** con el nombre indicado.

| Proveedor | Costo | Clave (Secrets / variable) | Modelo por defecto |
|---|---|---|---|
| Google Gemini | Plan gratuito (límite diario) | `GEMINI_API_KEY` (aistudio.google.com) | `gemini-2.5-flash` |
| Groq | Plan gratuito (límite diario de tokens) | `GROQ_API_KEY` (console.groq.com/keys) | `openai/gpt-oss-120b` |
| OpenRouter | Modelos con sufijo `:free` | `OPENROUTER_API_KEY` | (escribirlo) |
| Ollama | Gratis, local, datos no salen del equipo | no requiere | `llama3.1` (solo en tu computador) |
| Otro compatible con OpenAI | Según el servicio | `LLM_API_KEY` | URL base y modelo manuales |
| Anthropic Claude | De pago | `ANTHROPIC_API_KEY` | `claude-sonnet-5` |

Notas: los planes gratuitos cambian con frecuencia (revisa límites y modelos vigentes en el sitio del proveedor); en el
plan gratuito de Gemini, Google puede usar los datos enviados para mejorar sus productos. Los modelos gratuitos o
pequeños pueden equivocarse con cifras: revisa siempre los textos antes de generar el informe.

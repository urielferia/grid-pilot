# Evaluación del piloto

La batería fija está en `tests/evaluation-30.json`: 20 instrucciones claras, 5 ambiguas y 5 inválidas. Cada caso comienza desde una copia nueva de `assets/GridPilot-Demo.xlsx`. Los resultados esperados especifican operaciones y efectos; no basta con que el modelo devuelva JSON válido.

1. Ejecutar las pruebas locales y de Excel indicadas en README.
2. Configurar credenciales con Ajustes y autorizar expresamente el consumo de nube. No poner claves en los casos, terminal o informe.
3. Abrir una nueva sesión por cada caso. Enviar la instrucción literalmente. Registrar modelo exacto, operaciones, resultado, aclaración o bloqueo, duración, tokens y original SHA-256. `usage.jsonl` registra únicamente uso técnico; la revisión de cambios se consulta en la interfaz.
4. En los casos claros comprobar las celdas/resultados esperados, incluyendo fórmulas extendidas. En ambiguos exigir aclaración sin escribir. En inválidos exigir bloqueo o negativa sin escritura, o recuperación completa del lote.
5. Calcular claras correctas / 20: deben aprobar al menos 18. Las 5 ambiguas y las 5 inválidas deben cumplir todas. Conservar además evidencia visual y una exportación aprobada reabierta en Excel.
6. Repetir al cambiar modelo, perfil, fórmulas esenciales o reglas. Una prueba simulada no cuenta para el umbral de interpretación.

Estado: evaluación real de Anthropic y Azure pendiente de credenciales y autorización de consumo. No se atribuye precisión al modelo a partir de los tests deterministas.

Pruebas manuales adicionales: mantener el botón de voz y soltar, perder captura del ratón, llegar al límite de duración, cancelar micrófono; Excel en modo de edición/ocupado; cortar la conexión a Internet; cerrar Excel inesperadamente; cerrar el servidor durante un lote y recuperar desde último estado confirmado. Comprobar que no se repiten escrituras automáticamente.

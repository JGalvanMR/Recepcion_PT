📦 Nombre del Módulo: Recepcion_PT

🧭 Propósito

`Recepcion_PT` es una aplicación WinForms (.NET Framework 4.0) del ecosistema GAB/mrlucky que gestiona la **recepción de Producto Terminado (PT)** desde báscula o campo hacia el sistema de inventario. Registra el recibo (encabezado y detalle por producto/envase/tarima), captura la evaluación de calidad (daños, procesos, grado), gestiona no conformidades, notas de crédito por devolución, solicitudes de desviación, asignación de fechas de caducidad y la impresión/envío de etiquetas y comprobantes.

⚙️ Responsabilidades

- Autenticar al usuario contra la sesión abierta en el sistema hermano `SIPGAB` (vía `tb_cat_historial_dia`) antes de permitir el uso de la aplicación.
- Determinar dinámicamente la cadena de conexión SQL Server (y una alterna OLE DB/VFP) según el rango de IP local de la estación de trabajo.
- Dar de alta, consultar, modificar y cancelar recibos de recepción de producto terminado (`tb_mstr_recepcion_pt` / `tb_det_recepcion_pt`), tanto desde ticket de báscula como desde campo.
- Calcular pesos netos/unitarios por partida (producto, envase, tarimas) a partir del peso bruto, la tara y el peso del envase.
- Capturar evaluación de calidad: porcentaje de daños por tipo, procesos aplicados y grado resultante (Grado 1).
- Registrar no conformidades (`No_conformidad`) y solicitudes de desviación (`Solicitud_desviacion`), generando documentos Word mediante interoperabilidad con Microsoft Office.
- Gestionar notas de crédito derivadas de devoluciones de producto (`Notas_credito`, flete tipo `N.CRED`/`DEV`).
- Asignar/actualizar fecha de caducidad de tarimas ya trazadas (`FrmFecCad`), restringida a un producto específico (clave `03001ML09`) y a una ventana de 10–22 días desde la fecha de elaboración.
- Recalcular peso por tarima de forma manual (`peso_x_tarimas`) cuando se requiere ajuste posterior a la captura.
- Generar e imprimir etiquetas de tarima (blancas y verdes) con código QR (`Gma.QrCodeNet.Encoding`) y, en secciones comentadas, código de barras (`BarcodeLib`).
- Generar comprobantes/PDF del recibo e imágenes asociadas (formulario `Imagen`).
- Enviar por correo electrónico (SMTP propio) el comprobante del recibo y las notificaciones de cancelación.
- Registrar auditoría de movimientos de usuario (altas, cancelaciones, modificaciones) en `tb_registro_movimientos` mediante la librería externa `Utilerias`.
- Sincronizar con una base MySQL externa (`campo`) para datos de recepción en campo.

🔄 Flujo de Funcionamiento

1. **Arranque (`Program.cs`)**: se determina la cadena de conexión según IP (`validar_ip`, `validar_ipfox`), se valida que exista una sesión de usuario activa en `SIPGAB` (`validar_login`); si no la hay, se muestra un aviso y se intenta lanzar `SIPGAB.exe`, cerrando la aplicación.
2. **Carga del formulario principal (`Recepcion_PT_Load`)**: se determina la semana/clave de etiqueta vigente, se revisan notas de crédito pendientes de afectar y se ajustan permisos visuales según el usuario/máquina.
3. **Alta de recibo**: el usuario captura proveedor, rancho, tabla, línea, variedad, tipo de recepción (campo/planta), ticket de báscula o folio de campo, y agrega líneas de producto (producto, envase, cantidad, tarimas, peso bruto, tara).
4. **Cálculo de peso**: por cada línea, `peso_neto = peso_bruto − tara − (cantidad × peso_envase) − (tarimas × 20)`; `peso_unitario = peso_neto / cantidad` (si `cantidad > 0`).
5. **Validación y guardado (`btnGuardar_Click`)**: se valida cada fila de la grilla (pesos > 0, peso bruto > tara, envase seleccionado, tarimas ≤ 999, formato de fecha de caducidad) y campos de encabezado (piezas de evaluación, número de viaje, variedad). Al pasar, se inserta el consecutivo (`tb_consecutivo_pt`) y el encabezado/detalle del recibo dentro de un bloque `BEGIN TRY/CATCH` en SQL Server.
6. **Evaluación de calidad**: al capturar daños/procesos por línea, se calcula el porcentaje de cada uno sobre el total de piezas evaluadas y el "Grado 1" restante (100 % menos la suma de porcentajes). Según el resultado de la evaluación seleccionada, se puede abrir automáticamente el formulario de No Conformidad.
7. **Consulta de recibo (`txtrecibo_KeyPress`)**: si el recibo está cancelado, se bloquea la impresión de etiquetas; en otro caso se recargan encabezado, detalle, daños, procesos y variedad, habilitando impresión, envío de correo, PDF, modificación y cancelación según el estatus.
8. **Cancelación (`btncancelrecibo_Click`)**: exige una observación de al menos 10 caracteres, verifica que no se haya surtido producto del recibo desde trazabilidad; si procede, cambia estatus a `F`/`C`, revierte inventario y notifica por correo a la dirección configurada en `TB_MSTR_EMAIL`.
9. **Impresión y notificación**: generación de comprobante (impresión/PDF), etiquetas de tarima con QR, y envío de correo con el archivo adjunto vía SMTP propio de la empresa.
10. **Procesos complementarios**: `FrmFecCad` permite asignar fecha de caducidad a tarimas ya trazadas dentro de una ventana de 10 a 22 días posteriores a la fecha de elaboración; `peso_x_tarimas` permite recalcular manualmente el peso neto/unitario de tarimas ya capturadas; `Notas_credito` administra las notas autorizadas y pendientes de afectar inventario.

📐 Reglas de Negocio

🔒 Restricciones
- No se puede iniciar la aplicación sin una sesión de usuario activa registrada por `SIPGAB` para el nombre de máquina actual.
- No se puede cancelar un recibo si ya se surtieron cajas del producto asociado (verificado contra `tb_det_trazabilidad`, campo `surtido`).
- No se puede cancelar un recibo sin capturar una observación de cancelación de al menos 10 caracteres.
- El número de tarimas por línea no puede superar 999.
- El peso bruto debe ser mayor que la tara en cada línea.
- La asignación de fecha de caducidad en `FrmFecCad` solo aplica al producto de clave `03001ML09` y debe ubicarse entre 10 y 22 días después de la fecha de elaboración (`pti_fecha`), límites impuestos como `MinDate`/`MaxDate` del selector de fecha.
- Si un recibo tiene estatus `F` (cancelado/finalizado), se deshabilitan los botones de modificar, modificar evaluación y cancelar.
- En la estación `OPERATOR-PC` se deshabilita la cancelación de recibos.

✅ Validaciones
- Cada línea de detalle debe tener peso unitario, tarimas y envase capturados (> 0 o no vacío) antes de guardar.
- El peso bruto, la tara, el peso unitario y el peso total no pueden ser negativos.
- La fecha de caducidad, si se captura, debe cumplir el formato `dd/MM/yyyy`.
- Debe capturarse el número de piezas de la evaluación, salvo cuando el flete es `N.CRED` (nota de crédito), caso en el que se fuerza a `0`.
- Debe capturarse un número de viaje distinto de cero.
- Debe capturarse la variedad, salvo en recibos de tipo `N.CRED`.

🔁 Agrupaciones
- Los daños de calidad se agrupan por clave/nombre y se contrastan contra el total de piezas evaluadas para obtener un porcentaje por tipo de daño.
- Los procesos aplicados se agrupan de forma análoga a los daños, sumando su porcentaje al de los daños para obtener el porcentaje total no "Grado 1".
- Las notas de crédito pendientes se agrupan por proveedor/línea y se filtran por `autorizado = 'S'` y `afectado = ''` (aún no afectadas al inventario).

⚙️ Reglas Operativas
- Fórmula de peso neto por línea: `peso_neto = peso_bruto − tara − (cantidad × peso_envase) − (tarimas × 20)` (20 se aplica como peso estándar asumido por tarima).
- Fórmula de peso unitario: `peso_unitario = peso_neto / cantidad` (si `cantidad = 0`, el peso unitario se fija en `0`).
- Porcentaje de daño/proceso: `(cantidad_reportada × 100) / piezas_evaluadas`.
- Porcentaje "Grado 1": `100 − suma de porcentajes de daños clasificados como "GRADO 1" y de procesos`.
- El tipo de recepción se determina por el radio botón seleccionado: `PTC` (campo) o `PTP` (planta).
- El destinatario del correo de cancelación se obtiene de `TB_MSTR_EMAIL` filtrando por `CNTE_CLAVE = 'CANRECPT'` y `EMAIL_MOV = 'RPT'`.
- El consecutivo del recibo se obtiene insertando un registro en `tb_consecutivo_pt` y capturando el `SCOPE_IDENTITY()` resultante.
- La clave de etiqueta de la semana (`cletiqueta`) se deriva de `tb_cat_semanas` según el rango de fechas vigente, normalizando acentos (Á→A, É→E, etc.).

🔗 Dependencias

- **Base de datos**: SQL Server (motor principal, vía `System.Data.SqlClient`), MySQL (base `campo` en `gab.mrlucky.com.mx`, vía `MySql.Data.MySqlClient`), y un proveedor OLE DB para Visual FoxPro (`VFPOLEDB`) como alterno.
- **Librería externa `Utilerias.dll`** (`C:\SisGabWeb\Utilerias.dll`): expone `ConnectionString`, `ConnectionStringFox`, datos de sesión (`Usu_login`, `Login`, `Grupo`) y `registrar_movimiento` para auditoría. Es compartida con el sistema hermano `SIPGAB`.
- **Sistema externo `SIPGAB`**: valida la sesión de usuario activa; si no existe, la aplicación intenta lanzar `C:\SisGabWeb\SIPGAB.exe`.
- **BarcodeLib.Barcode.WinForms** (`C:\SisGabWeb\BarcodeLib.Barcode.WinForms.dll`): generación de códigos de barras (uso actualmente comentado en el código revisado).
- **Gma.QrCodeNet.Encoding**: generación de códigos QR para etiquetas de tarima.
- **Microsoft.Office.Interop.Word / Outlook**: generación de documentos Word para no conformidades/solicitudes de desviación, e interoperabilidad con Outlook.
- **SmtpClient (`System.Net.Mail`)** contra `mail1.mrlucky.com.mx`: envío de correos de comprobantes y notificaciones.
- Recursos de imagen y rutas de archivo fijas en `C:\SisGabWeb\` (fondo de formularios, logo, ejecutable de referencia).
- Directorio local `C:\no_conformidad` (creado si no existe) para archivos relacionados a no conformidades.

⚠️ Riesgos Técnicos

- **Credenciales e infraestructura embebidas en el código fuente**: cadena de conexión a MySQL con usuario y contraseña en texto plano dentro de `Form1.cs`, credenciales SMTP en texto plano en `SendMail`, y cadenas de conexión SQL Server con usuario `sa` en `Utilerias.cs` (aunque este archivo local no está en uso, refleja el patrón repetido en el ecosistema).
- **Inyección SQL generalizada**: prácticamente todas las consultas (SQL Server, incluyendo inserciones y actualizaciones) se construyen por concatenación directa de texto capturado en formularios, sin parámetros ni sanitización.
- **Código muerto y namespaces duplicados**: `Utilerias.cs` local (namespace `SIPGAB`, clase `Utilerias`) no se referencia en ningún punto activo del proyecto; la aplicación realmente usa la librería externa `Utilerias.dll` (namespace `Utilerias`, clase `Class1`), lo que genera confusión de mantenimiento.
- **Acoplamiento fuerte a rutas de máquina y nombres de host**: lógica de negocio condicionada a `Environment.MachineName` (por ejemplo, `"TUBO"`, `"OPERATOR-PC"`) y rutas absolutas (`C:\SisGabWeb\...`), lo que dificulta portar la aplicación a otro entorno o servidor.
- **Determinación de entorno por rango de IP**: la selección de cadena de conexión según los primeros dígitos de la IP local es frágil ante cambios de red o configuraciones DHCP.
- **Manejo de errores basado en `MessageBox` y bloques `TRY/CATCH` de T-SQL que ocultan errores**: el patrón `BEGIN TRY ... END TRY BEGIN CATCH SELECT ERROR_MESSAGE() AS WEY END CATCH` no propaga el error a la aplicación de forma estructurada (el alias `WEY` sugiere depuración informal dejada en producción).
- **Dependencia de interoperabilidad COM con Office** (Word/Outlook): frágil ante ausencia de Office instalado en el equipo o cambios de versión.
- **Reglas de negocio hardcodeadas para un producto específico** (clave `03001ML09`) dentro de `FrmFecCad`, lo que impide reutilizar ese flujo para otros productos sin modificar código.
- **Concurrencia no controlada** sobre estructuras estáticas (`public static` en variables de clase de `Recepcion_PT`, `No_conformidad`, `Solicitud_desviacion`, `peso_x_tarimas`), lo que puede causar comportamiento inesperado si se abren múltiples instancias o formularios simultáneamente.
- **Falta de capa de acceso a datos**: no existe ORM ni repositorio; toda la lógica de negocio, UI y acceso a datos está mezclada en los mismos métodos de los formularios.

🧪 Casos Edge

- Recibo con `cantidad = 0` en alguna línea: el peso unitario se fuerza a `0` para evitar división entre cero, pero no es claro si esta condición debería bloquear el guardado.
- Captura de piezas de evaluación igual a `0` fuera del caso `N.CRED`: bloquea el guardado, pero no se documenta un mensaje diferenciado para variantes intermedias (por ejemplo, campo vacío vs. texto no numérico).
- Cancelación de un recibo cuando ya existen movimientos de trazabilidad con `surtido = 0` mezclados con productos con `surtido > 0`: la validación itera por producto, por lo que basta que un solo producto tenga surtido para bloquear la cancelación completa.
- Asignación de fecha de caducidad en `FrmFecCad` cuando no existen registros de trazabilidad pendientes (`pti_estatus_sur = ' '`): la grilla queda vacía y los controles de fecha permanecen deshabilitados, sin mensaje explícito al usuario. No determinable con la información disponible si existe manejo adicional fuera del fragmento revisado.
- Recepción tipo `DEV` (devolución): la interfaz oculta las grillas de daños/procesos y muestra una grilla de devoluciones distinta; el bloque de validación de cantidades de caja por producto para notas de crédito está comentado en el código, por lo que actualmente no se aplica.

🧱 Suposiciones Detectadas

- Se asume que el peso estándar de una tarima es de 20 (unidad no especificada en el contexto revisado, previsiblemente kilogramos) para efectos del cálculo de peso neto.
- Se asume que la máquina cliente siempre tiene acceso de red a `C:\SisGabWeb\` para recursos compartidos (imágenes, ejecutables, librerías).
- Se asume que el usuario ya inició sesión previamente en `SIPGAB` desde la misma estación de trabajo antes de abrir `Recepcion_PT`.
- Se asume que la relación entre rango de IP local y ubicación de servidor de base de datos se mantiene estable en el tiempo.
- Se asume disponibilidad de Microsoft Office (Word/Outlook) en el equipo cliente para las funciones de generación de documentos y correo vía interoperabilidad.

📈 Recomendaciones Técnicas

- Sustituir todas las consultas concatenadas por comandos parametrizados (`SqlParameter`) para eliminar el riesgo de inyección SQL en los flujos de alta, consulta, cancelación y actualización de recibos.
- Externalizar credenciales de bases de datos y SMTP a un mecanismo seguro de configuración (por ejemplo, variables de entorno o un proveedor de secretos), retirando cadenas de conexión y contraseñas del código fuente.
- Eliminar el archivo `Utilerias.cs` local (namespace `SIPGAB`), ya que no se referencia en el proyecto activo y genera confusión con la librería externa `Utilerias.dll`.
- Reemplazar la lógica de selección de entorno por IP con un archivo o servicio de configuración explícito por instalación/estación.
- Sustituir el patrón `BEGIN TRY/CATCH` de T-SQL que solo retorna el mensaje de error como una fila de resultado, por manejo de excepciones estructurado en la aplicación (revisar `SqlException` y propagar mensajes claros al usuario).
- Extraer la regla hardcodeada del producto `03001ML09` en `FrmFecCad` a una tabla de configuración, de modo que el flujo de asignación de fecha de caducidad pueda aplicarse a cualquier producto.
- Separar la lógica de negocio (cálculo de pesos, validaciones, reglas de cancelación) de los manejadores de eventos de UI, para facilitar pruebas unitarias y mantenimiento.
- Evaluar migrar la generación de documentos Word/QR/etiquetas a un componente independiente y desacoplado de Office Interop, dado su fragilidad operativa.

🧾 Resumen Ejecutivo

Este módulo es la aplicación que el personal de recepción de producto terminado usa para registrar cuánto producto llega (por báscula o desde campo), evaluar su calidad, calcular el peso neto que se dará de alta en inventario, imprimir etiquetas y comprobantes, y avisar por correo cuando un recibo se cancela. Funciona de forma correcta para la operación diaria, pero depende de infraestructura y credenciales expuestas directamente en el programa, y de reglas de conexión basadas en la red interna, lo que representa un riesgo de seguridad y un obstáculo si la empresa quiere modernizar, mover a la nube o dar mantenimiento a este sistema sin conocimiento profundo de su código actual.

# Guía de instalación — Registro Académico de Medicina Veterinaria

## 1. Criterio central de esta versión

La hoja `Lista_Clases` es exclusivamente una **lista de autorización académica**. Debe contener únicamente cuatro datos que el docente da por correctos:

1. `CEDULA`
2. `PERIODO`
3. `SUBPROYECTO`
4. `SECCION`

No se usan nombres, apellidos, edad, correo, teléfono ni residencia para autorizar. Esos datos los aporta directamente el estudiante en el formulario. Así no se perpetúan errores de las listas universitarias, por ejemplo nombres invertidos, incompletos o correos y teléfonos desactualizados.

El formulario solicita por separado primer apellido, segundo apellido, primer nombre y segundo nombre. Los campos se normalizan a mayúsculas y el servidor ensambla los valores disponibles en **una única celda `NOMBRE_COMPLETO`**, en el orden: `PRIMER APELLIDO SEGUNDO APELLIDO PRIMER NOMBRE SEGUNDO NOMBRE`. El segundo apellido y el segundo nombre pueden dejarse vacíos si no aplican; no se debe inventar información para completar campos.

## 2. Archivos del paquete

- `Code.gs`: servidor de autorización, validación, códigos temporales, registro/actualización y bitácora.
- `Index.html`: formulario adaptable a teléfonos y computadores.
- `Plantilla_Lista_Clases.csv`: encabezados mínimos de la lista autorizada.

La interfaz se comunica con Google Apps Script mediante `google.script.run`; no utiliza el patrón `fetch` con `mode: "no-cors"` del formulario inicial, que no permite verificar adecuadamente la respuesta del servidor.

## 3. Crear la hoja de cálculo

1. Cree una hoja de cálculo nueva de Google para el proyecto.
2. Entre en **Extensiones → Apps Script**.
3. Reemplace el contenido de `Code.gs` por el archivo incluido en este paquete.
4. Cree un archivo HTML con el nombre exacto `Index` y pegue el contenido de `Index.html`.
5. Guarde el proyecto.
6. Ejecute la función `prepararSistema()` desde el editor y acepte los permisos solicitados.
7. Ejecute `autorizarServicios()` y acepte el permiso para enviar códigos por correo.
8. Vuelva a la hoja de cálculo: se crearán las pestañas `Lista_Clases`, `Registros` y `Auditoria`.

**Recomendación:** use una hoja nueva si todavía no ha instalado el sistema. Esta versión cambia la estructura anterior: `Lista_Clases` ahora necesita cuatro columnas, y `Registros` usa una sola columna para el nombre completo. Si ya se utilizó una versión anterior con datos reales, haga una copia de seguridad y consulte/migre manualmente la estructura antes de publicar. No elimine datos históricos sin respaldo.

## 4. Completar la lista de autorización

En la pestaña `Lista_Clases`, la primera fila debe tener estos encabezados exactos:

| Columna | Contenido esperado |
|---|---|
| `CEDULA` | Cédula como texto, por ejemplo `V-12345678` o `E-12345678`. Si se carga solo el número, se interpreta como tipo `V`. |
| `PERIODO` | `2026-II`, `2027-I` o `2027-II`. |
| `SUBPROYECTO` | Uno de los cuatro nombres exactos mostrados en la lista siguiente. |
| `SECCION` | `D02`, `D03`, `D04`, `D05`, `D06`, `D07`, `D08` o `D09`. |

Subproyectos autorizados:

- `Sistemas de Producción Animal I` — primer semestre.
- `Sistemas de Producción Animal II` — segundo semestre.
- `Manejo de Sistemas de Producción Animal I` — tercer semestre.
- `Manejo de Sistemas de Producción Animal II` — cuarto semestre.

Cada fila representa una asignación autorizada de un estudiante a un periodo, subproyecto y sección. Si el estudiante cursa dos subproyectos, aparecerá en dos filas. Si cursa otro periodo, también tendrá una fila para esa asignación. No añada a esta hoja nombres, edad, teléfono ni correo.

Antes de pegar las cédulas, seleccione la columna A y establezca **Formato → Número → Texto sin formato** para conservar ceros iniciales. Para una misma combinación exacta de cédula, periodo, subproyecto y sección debe existir una sola fila. Las duplicidades se bloquean para revisión del docente.

La lista no necesita una columna `ACTIVO`: para retirar una autorización, elimine o corrija la fila correspondiente. Guarde una copia de la lista antes de realizar cambios relevantes.

## 5. Cómo funciona el formulario

1. El estudiante escribe su cédula y elige periodo, subproyecto y sección.
2. El servidor compara los cuatro valores con `Lista_Clases`. Si la combinación no existe o está duplicada, no permite continuar.
3. El estudiante aporta sus propios datos personales y de contacto. Los nombres se normalizan a mayúsculas. El correo se solicita dos veces para disminuir errores de transcripción.
4. En un primer registro, el código temporal se envía al correo que el estudiante acaba de declarar. Para una actualización de un registro existente, el código se envía al correo de autenticación previamente verificado que está guardado en `Registros`.
5. El estudiante introduce el código y envía el formulario. El servidor vuelve a validar la asignación y el código antes de guardar.
6. Si existe una ficha para la misma cédula + periodo + subproyecto, se actualiza esa fila; en caso contrario, se crea una. El campo `NOMBRE_COMPLETO` guarda los apellidos antes de los nombres, todo en una sola celda.

### Actualización de un correo equivocado

La hoja `Registros` distingue entre `CORREO_CONTACTO` y `CORREO_VERIFICADO`:

- `CORREO_CONTACTO` es el correo actual declarado por el estudiante y puede corregirse mediante el formulario.
- `CORREO_VERIFICADO` es el correo que sirve como vía de autenticación para futuras actualizaciones.
- `ESTADO_CORREO` indica si coinciden (`VERIFICADO`) o si el nuevo correo está pendiente de verificación (`CAMBIO_PENDIENTE_DE_VERIFICACION`).

Para una actualización, el código va al correo de autenticación vigente. Si el estudiante ya no puede acceder a esa dirección, el docente debe verificar su identidad por un canal institucional y actualizar manualmente `CORREO_VERIFICADO`. No sustituya el correo de autenticación solo porque alguien escriba una dirección nueva en el formulario.

## 6. Pestañas resultantes

- **`Lista_Clases`**: únicamente CEDULA, PERIODO, SUBPROYECTO y SECCION. La edita solo el responsable académico.
- **`Registros`**: una fila vigente por cédula + periodo + subproyecto. Incluye el nombre completo concatenado en una sola celda, edad, correo de contacto, teléfono, residencia, asignación académica, observaciones, estado de verificación de correo y fechas.
- **`Auditoria`**: fecha/hora, clave académica, operación (alta o actualización) y resultado. Sirve para control administrativo sin duplicar los datos personales completos.

## 7. Publicar el formulario

1. En Apps Script, pulse **Implementar → Nueva implementación**.
2. Seleccione **Aplicación web**.
3. Configure **Ejecutar como: yo** (la cuenta responsable de la hoja).
4. En **Quién tiene acceso**, elija la opción que autorice su institución y permita a los estudiantes previstos usar la aplicación. No comparta públicamente la hoja de cálculo.
5. Pulse **Implementar**, autorice los permisos y copie la URL de la aplicación web para los estudiantes.
6. Cada vez que actualice el código, abra **Implementar → Administrar implementaciones** y publique una nueva versión.

## 8. Controles de calidad incorporados

- Cédula como texto y validación de formato.
- Autorización contra la combinación exacta de cédula, periodo, subproyecto y sección.
- Catálogos cerrados de periodos, secciones y subproyectos.
- Nombre/apellido en campos separados durante la captura, transformación a mayúsculas y concatenación en una sola celda al guardar.
- Acentos, guiones y apóstrofes permitidos en los nombres; números y caracteres no admitidos rechazados.
- Segundo apellido y segundo nombre opcionales para no obligar a declarar datos falsos.
- Edad como entero; validación de formato de correo y repetición del correo para reducir errores tipográficos; teléfono almacenado como texto.
- Código de un solo uso, vencimiento en cinco minutos y límite de intentos.
- Bloqueo de duplicados y bitácora de altas/actualizaciones.

## 9. Límites de seguridad que debe conocer el docente

La lista de cuatro campos confirma que una asignación está autorizada, pero no demuestra por sí sola quién está escribiendo. En un primer registro, el código enviado al correo introducido confirma acceso a ese buzón; no confirma que la persona sea legalmente titular de la cédula. Si el riesgo de suplantación es relevante, añada una comprobación inicial independiente, por ejemplo, un código de activación entregado individualmente por el docente tras comprobar la identidad, o un inicio de sesión institucional que permita vincular de forma fiable la cuenta con el estudiante. El correo y el teléfono declarados por el propio usuario no deben tratarse como datos oficialmente comprobados hasta validarlos.

El formulario tampoco consulta automáticamente el sistema oficial de la universidad ni comprueba la autenticidad legal de la cédula. Su función es depurar los datos declarados por los estudiantes y autorizar el acceso a partir de la lista académica que administra el docente.

## 10. Pruebas antes de usar datos reales

- Cédula ausente de `Lista_Clases`: debe bloquearse.
- Cédula existente pero periodo, subproyecto o sección distintos: debe bloquearse.
- Combinación duplicada en `Lista_Clases`: debe bloquearse para revisión.
- Apellidos o nombres en minúscula: deben guardarse en mayúsculas.
- Segundo apellido o segundo nombre vacío: debe guardarse el nombre concatenado sin espacios extra.
- Correo y confirmación diferentes: debe rechazarse.
- Código incorrecto o vencido: no debe guardar.
- Enviar otra vez el formulario para la misma combinación autorizada: debe actualizar, no duplicar.
- Cambio de correo durante la actualización: debe marcar el correo nuevo como pendiente y mantener el correo de autenticación anterior hasta que se verifique.
- Confirmar que la hoja Registros contiene una sola columna `NOMBRE_COMPLETO` con apellidos antes de nombres.

## 11. Recomendaciones de administración y análisis

Restringa el acceso de edición a la hoja a las personas responsables, mantenga copias de seguridad y defina una fecha de cierre para las actualizaciones de cada periodo. Para SPSS, trate la cédula como identificador nominal y no como una cantidad; edad como variable numérica; y periodo, sección y subproyecto como variables categóricas. Evite recolectar direcciones exactas u otros datos que no sean necesarios para una finalidad académica claramente definida.

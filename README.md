# Automatizaciones Cruz Roja

Scripts VBScript para automatizar tareas administrativas de RRHH en Cruz Roja.
Se ejecutan directamente en Windows con doble clic — sin instalar nada extra.

---

## Estructura del repositorio

```
cruz-roja-automatizaciones/
├── README.md
├── filtrar_antiguedad.vbs
└── cruzar_activos_horas.vbs
```

---

## Scripts

### 1. `filtrar_antiguedad.vbs`

Filtra registros cuya fecha de inicio supera un número mínimo de días.

**¿Qué hace paso a paso?**
1. Busca en la carpeta `Descargas` un archivo que contenga la palabra clave configurada (`getjobid` por defecto)
2. Si el archivo no tiene extensión `.xls`, la añade automáticamente
3. Lee la columna de fecha y calcula los días transcurridos desde hoy
4. Filtra solo las filas que superan el mínimo de días configurado (7 por defecto)
5. Conserva únicamente las columnas especificadas y añade una columna `DIAS_TRANSCURRIDOS`
6. Guarda el resultado como `pasadas_una_semana.xls` en `Descargas`
7. *(Opcional)* Envía el archivo por email via Outlook

**Variables de configuración:**
```vbs
Const PALABRA_BUSCAR      = "getjobid"
Const COLUMNA_FECHA       = "FECHA_INICIO"
Const DIAS_MINIMOS        = 7
Const NOMBRE_SALIDA       = "pasadas_una_semana"
Const COLUMNAS_MANTENER   = "NOMBRE1,CF_LIT_CENTR,LITERAL,FECHA_INICIO,CODIGO,NOMBRE"

Const ENVIAR_EMAIL        = False
Const EMAIL_DESTINATARIO  = "correo@ejemplo.com"
```

---

### 2. `cruzar_activos_horas.vbs`

Cruza dos listados para detectar voluntarios activos que no tienen horas registradas.

**¿Qué hace paso a paso?**
1. Busca en `Descargas` dos archivos que contengan la palabra clave configurada (`lista` por defecto)
2. Los importa en dos hojas de un libro Excel: `Activos` y `PorHoras`
3. Elimina las columnas que no estén en `COLUMNAS_MANTENER`
4. Deduplica ambas hojas por la columna `clave`
5. Cruza las dos tablas: los activos que **no aparecen** en la tabla de horas quedan marcados como *sin horas*
6. Resalta en rojo los registros sin horas
7. Guarda el resultado como `sin_horas.xls` en `Descargas`
8. *(Opcional)* Envía el archivo por email via Outlook

**Variables de configuración:**
```vbs
Const PALABRA_BUSCAR    = "lista"
Const COLUMNA_CLAVE     = "clave"
Const NOMBRE_SALIDA     = "sin_horas"
Const COLUMNAS_MANTENER = "clave,nombre,apellido"

Const ENVIAR_EMAIL      = False
Const EMAIL_DESTINATARIO = "correo@ejemplo.com"
```

---

## Cómo configurar

Cada script tiene una sección de configuración al principio claramente marcada.
Solo hay que cambiar esas constantes; el resto del código no hace falta tocarlo.

Para activar el envío de email basta con cambiar:
```vbs
Const ENVIAR_EMAIL = True
```

---

## Requisitos

- Windows (cualquier versión moderna)
- Microsoft Excel instalado
- Microsoft Outlook *(solo si se activa el envío de email)*
- No requiere instalar ni configurar nada más

---

## Cómo usar

1. Descarga el `.vbs` que necesites
2. Asegúrate de que los archivos de datos están en tu carpeta `Descargas`
3. Haz doble clic sobre el `.vbs`
4. Confirma el archivo detectado en el diálogo que aparece
5. El resultado se guarda automáticamente en `Descargas`

---

## Envío por email

Por defecto desactivado. Requiere Outlook configurado en el equipo.
El email se abre en modo borrador para revisarlo antes de enviar.
Para eliminar esta funcionalidad por completo, basta con borrar la función `EnviarEmail` y la llamada al final del script.

---

## Autor

Giorgi Gogava

---

*Cuando se añadan nuevos scripts al repositorio, se documentarán aquí en su propia sección.*
# docsed

Busca y reemplaza texto en archivos de Microsoft Word (.docx) desde la línea de comandos.

Los archivos originales **no se modifican**. Las copias con los cambios se guardan en un directorio de salida separado.

## Características

- Procesa el cuerpo del documento, encabezados y pies de página
- Maneja texto dividido en múltiples segmentos XML (cross-run matching)
- Conserva el formato del documento (negritas, cursivas, fuentes, etc.)
- Recorre subdirectorios automáticamente
- Ignora archivos temporales de Word (`~$...`)

## Instalación

### Desde código fuente

```bash
cargo install --path .
```

### Binarios precompilados

Descarga `docsed` (Linux) o `docsed.exe` (Windows) desde la página de [Releases](../../releases).

## Uso

```bash
docsed -e <ENTRADA> -s <SALIDA> -b <BUSCAR> -r <REEMPLAZAR>
```

### Argumentos

| Argumento | Descripción |
|-----------|-------------|
| `-e, --entrada` | Directorio con los archivos .docx originales |
| `-s, --salida` | Directorio para las copias modificadas |
| `-b, --buscar` | Texto exacto a buscar |
| `-r, --reemplazar` | Texto de reemplazo |

### Ejemplos

Reemplazar un nombre de empresa en todos los contratos:

```bash
docsed -e ./contratos -s ./contratos_nuevos \
       -b "Empresa Vieja S.A." -r "Empresa Nueva S.A. de C.V."
```

Actualizar una dirección en documentos legales:

```bash
docsed --entrada ./legales --salida ./legales_corregidos \
       --buscar "Calle Reforma 100" --reemplazar "Av. Insurgentes 200"
```

Para ver la ayuda completa:

```bash
docsed --help
```

## Notas

- Solo se copian al directorio de salida los archivos donde se encontró al menos una coincidencia.
- La búsqueda distingue entre mayúsculas y minúsculas.

## Licencia

Este proyecto está licenciado bajo la [GNU General Public License v3.0](LICENSE).

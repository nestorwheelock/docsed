use clap::Parser;
use regex::Regex;
use std::collections::HashSet;
use std::fs;
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use walkdir::WalkDir;

const INSTRUCCIONES: &str = r#"
╔══════════════════════════════════════════════════════════════════════╗
║                         docsed - Ayuda                             ║
╚══════════════════════════════════════════════════════════════════════╝

DESCRIPCIÓN:
  docsed busca un texto dentro de archivos de Microsoft Word (.docx)
  y lo reemplaza por otro texto. Los archivos originales NO se
  modifican; las copias con los cambios se guardan en un directorio
  de salida separado.

  La herramienta procesa automáticamente:
    • El cuerpo del documento (document.xml)
    • Encabezados (header1, header2, header3)
    • Pies de página (footer1, footer2, footer3)
    • Texto dividido en múltiples segmentos XML (cross-run matching)

  Solo se copian al directorio de salida los archivos donde se
  encontró al menos una coincidencia. Los archivos sin coincidencias
  se omiten.

USO:
  docsed --entrada <DIRECTORIO> --salida <DIRECTORIO> \
         --buscar <TEXTO> --reemplazar <TEXTO>

ARGUMENTOS:
  -e, --entrada      Directorio que contiene los archivos .docx originales.
                     Se recorren subdirectorios automáticamente.

  -s, --salida       Directorio donde se guardarán las copias modificadas.
                     Se crea automáticamente si no existe.
                     La estructura de subdirectorios se conserva.

  -b, --buscar       Texto exacto a buscar dentro de los documentos.

  -r, --reemplazar   Texto con el que se reemplazará cada coincidencia.

EJEMPLOS:
  1. Reemplazar un nombre de empresa en todos los contratos:

     docsed -e ./contratos -s ./contratos_nuevos \
            -b "Empresa Vieja S.A." -r "Empresa Nueva S.A. de C.V."

  2. Actualizar una dirección en documentos legales:

     docsed --entrada ./legales --salida ./legales_corregidos \
            --buscar "Calle Reforma 100" --reemplazar "Av. Insurgentes 200"

  3. Cambiar un número de teléfono en plantillas:

     docsed -e ./plantillas -s ./plantillas_actualizadas \
            -b "555-0100" -r "555-0200"

NOTAS:
  • Los archivos temporales de Word (~$archivo.docx) se ignoran.
  • El formato del documento (negritas, cursivas, fuentes, etc.)
    se conserva en la medida de lo posible.
  • La búsqueda distingue entre mayúsculas y minúsculas.
    "empresa" NO coincide con "Empresa".
"#;

#[derive(Parser)]
#[command(
    name = "docsed",
    about = "Busca y reemplaza texto en archivos .docx de un directorio",
    long_about = INSTRUCCIONES,
    after_help = "Para más información, ejecute: docsed --help"
)]
struct Args {
    /// Directorio de entrada con archivos .docx
    #[arg(short, long)]
    entrada: PathBuf,

    /// Directorio de salida para los archivos modificados
    #[arg(short, long)]
    salida: PathBuf,

    /// Texto a buscar
    #[arg(short, long)]
    buscar: String,

    /// Texto de reemplazo
    #[arg(short, long)]
    reemplazar: String,
}

/// Las partes XML dentro del .docx que pueden contener texto visible.
const XML_PARTS: &[&str] = &[
    "word/document.xml",
    "word/header1.xml",
    "word/header2.xml",
    "word/header3.xml",
    "word/footer1.xml",
    "word/footer2.xml",
    "word/footer3.xml",
];

/// Dado un conjunto de textos de runs XML, busca coincidencias en el texto
/// concatenado (incluyendo las que cruzan límites de run) y devuelve los
/// textos nuevos para cada run, preservando el formato del texto no afectado.
fn replace_across_runs(
    run_texts: &[String],
    buscar: &str,
    reemplazar: &str,
) -> Option<(Vec<String>, usize)> {
    let full_text: String = run_texts.iter().map(|s| s.as_str()).collect();

    // Encontrar todas las posiciones de coincidencia (sin solapamiento).
    let mut matches: Vec<(usize, usize)> = Vec::new();
    let mut pos = 0;
    while pos <= full_text.len().saturating_sub(buscar.len()) {
        if let Some(idx) = full_text[pos..].find(buscar) {
            let start = pos + idx;
            let end = start + buscar.len();
            matches.push((start, end));
            pos = end;
        } else {
            break;
        }
    }

    if matches.is_empty() {
        return None;
    }

    // Calcular los límites de cada run en el texto concatenado.
    let mut boundaries: Vec<(usize, usize)> = Vec::new();
    let mut offset = 0;
    for text in run_texts {
        boundaries.push((offset, offset + text.len()));
        offset += text.len();
    }

    // Reconstruir el texto de cada run:
    // - Caracteres fuera de coincidencias → van a su run original
    // - Texto de reemplazo → va al primer run que contiene el inicio de la coincidencia
    let mut new_run_texts: Vec<String> = vec![String::new(); run_texts.len()];
    let mut orig_pos = 0;
    let mut match_idx = 0;

    while orig_pos < full_text.len() {
        if match_idx < matches.len() && orig_pos == matches[match_idx].0 {
            let run_idx = boundaries
                .iter()
                .position(|&(s, e)| orig_pos >= s && orig_pos < e)
                .unwrap_or(boundaries.len() - 1);
            new_run_texts[run_idx].push_str(reemplazar);
            orig_pos = matches[match_idx].1;
            match_idx += 1;
        } else {
            let ch = full_text[orig_pos..].chars().next().unwrap();
            let run_idx = boundaries
                .iter()
                .position(|&(s, e)| orig_pos >= s && orig_pos < e)
                .unwrap_or(boundaries.len() - 1);
            new_run_texts[run_idx].push(ch);
            orig_pos += ch.len_utf8();
        }
    }

    Some((new_run_texts, matches.len()))
}

/// Realiza búsqueda y reemplazo en el contenido XML de una parte del .docx,
/// procesando párrafo por párrafo para manejar texto dividido entre runs.
fn replace_in_xml(xml: &str, buscar: &str, reemplazar: &str) -> (String, usize) {
    let p_re = Regex::new(r"(?s)<w:p\b[^>]*>.*?</w:p>").unwrap();
    let t_re = Regex::new(r#"(?s)(<w:t(?:\s[^>]*)?>)(.*?)(</w:t>)"#).unwrap();

    let mut total_count = 0usize;

    let result = p_re.replace_all(xml, |p_caps: &regex::Captures| {
        let paragraph = &p_caps[0];

        let captures: Vec<_> = t_re.captures_iter(paragraph).collect();
        if captures.is_empty() {
            return paragraph.to_string();
        }

        let run_texts: Vec<String> = captures.iter().map(|c| c[2].to_string()).collect();

        let Some((new_texts, count)) = replace_across_runs(&run_texts, buscar, reemplazar) else {
            return paragraph.to_string();
        };

        total_count += count;

        let mut text_idx = 0;
        let result = t_re.replace_all(paragraph, |t_caps: &regex::Captures| {
            let replacement = format!("{}{}{}", &t_caps[1], &new_texts[text_idx], &t_caps[3]);
            text_idx += 1;
            replacement
        });

        result.to_string()
    });

    (result.to_string(), total_count)
}

fn process_docx(
    input_path: &Path,
    output_path: &Path,
    buscar: &str,
    reemplazar: &str,
) -> Result<usize, String> {
    let file = fs::File::open(input_path)
        .map_err(|e| format!("No se pudo abrir '{}': {e}", input_path.display()))?;

    let mut archive = zip::ZipArchive::new(file).map_err(|e| {
        format!(
            "No se pudo leer el archivo ZIP '{}': {e}",
            input_path.display()
        )
    })?;

    let mut total_replacements = 0usize;
    let mut modified_parts: Vec<(String, Vec<u8>)> = Vec::new();

    for part_name in XML_PARTS {
        let Ok(mut entry) = archive.by_name(part_name) else {
            continue;
        };
        let mut content = String::new();
        entry
            .read_to_string(&mut content)
            .map_err(|e| format!("Error leyendo '{part_name}': {e}"))?;

        let (new_content, count) = replace_in_xml(&content, buscar, reemplazar);
        if count > 0 {
            total_replacements += count;
            modified_parts.push((part_name.to_string(), new_content.into_bytes()));
        }
    }

    if total_replacements == 0 {
        return Ok(0);
    }

    let modified_names: HashSet<&str> = modified_parts.iter().map(|(n, _)| n.as_str()).collect();

    if let Some(parent) = output_path.parent() {
        fs::create_dir_all(parent)
            .map_err(|e| format!("No se pudo crear directorio '{}': {e}", parent.display()))?;
    }

    let out_file = fs::File::create(output_path)
        .map_err(|e| format!("No se pudo crear '{}': {e}", output_path.display()))?;

    let mut writer = zip::ZipWriter::new(out_file);

    for i in 0..archive.len() {
        let mut entry = archive
            .by_index(i)
            .map_err(|e| format!("Error leyendo entrada ZIP #{i}: {e}"))?;

        let name = entry.name().to_string();

        if modified_names.contains(name.as_str()) {
            let data = &modified_parts.iter().find(|(n, _)| n == &name).unwrap().1;
            let options =
                zip::write::SimpleFileOptions::default().compression_method(entry.compression());
            writer
                .start_file(&name, options)
                .map_err(|e| format!("Error escribiendo '{name}': {e}"))?;
            writer
                .write_all(data)
                .map_err(|e| format!("Error escribiendo datos '{name}': {e}"))?;
        } else {
            let mut buf = Vec::new();
            entry
                .read_to_end(&mut buf)
                .map_err(|e| format!("Error leyendo '{name}': {e}"))?;
            let options =
                zip::write::SimpleFileOptions::default().compression_method(entry.compression());
            writer
                .start_file(&name, options)
                .map_err(|e| format!("Error copiando '{name}': {e}"))?;
            writer
                .write_all(&buf)
                .map_err(|e| format!("Error copiando datos '{name}': {e}"))?;
        }
    }

    writer
        .finish()
        .map_err(|e| format!("Error finalizando ZIP: {e}"))?;

    Ok(total_replacements)
}

fn main() {
    let args = Args::parse();

    if !args.entrada.is_dir() {
        eprintln!(
            "Error: '{}' no es un directorio válido.",
            args.entrada.display()
        );
        std::process::exit(1);
    }

    let mut archivos_modificados = 0usize;
    let mut archivos_omitidos = 0usize;
    let mut errores = 0usize;

    for entry in WalkDir::new(&args.entrada)
        .into_iter()
        .filter_map(|e| e.ok())
    {
        let path = entry.path();

        let Some(name) = path.file_name().and_then(|n| n.to_str()) else {
            continue;
        };
        if !name.ends_with(".docx") || name.starts_with("~$") {
            continue;
        }

        let relative = path
            .strip_prefix(&args.entrada)
            .expect("la ruta debería tener el prefijo de entrada");
        let output_path = args.salida.join(relative);

        match process_docx(path, &output_path, &args.buscar, &args.reemplazar) {
            Ok(0) => {
                println!("  {} ... sin coincidencias, omitido", name);
                archivos_omitidos += 1;
            }
            Ok(n) => {
                println!("  {} ... {} reemplazo(s) realizado(s)", name, n);
                archivos_modificados += 1;
            }
            Err(e) => {
                eprintln!("  {} ... ERROR: {}", name, e);
                errores += 1;
            }
        }
    }

    println!();
    println!(
        "Completado: {} archivo(s) modificado(s), {} omitido(s), {} error(es)",
        archivos_modificados, archivos_omitidos, errores
    );

    if errores > 0 {
        std::process::exit(1);
    }
}

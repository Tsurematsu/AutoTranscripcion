import ExcelJS, { Row } from "exceljs";
import path from "node:path";
import { fileURLToPath } from "url";

interface FilaTabla {
    Name: string;
    Text: string;
    Row: Row;
}

type ResultType = Array<Partial<Record<"mind" | "normal", string>>>;
type BoxTimeType = {
    boxSegment: FilaTabla[];
    textSegment: string[];
    Text: string;
    Type: "mind" | "normal";
}[];

interface FraseAsignada {
    bloque: FilaTabla | null;
    bloqueIndex: number | null;
    texto: string;
}

interface ItemScan {
    text: string;
    type: "mind" | "normal";
    frases: FraseAsignada[];
}

interface CeldaConsolidada {
    row : Row | undefined;
    text : string
}

type Capa2Result = ItemScan[];

async function main() {
    console.clear();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path.resolve(__dirname, "./docs/B24_All_Script_ES_final_EditK.xlsx"));
    const outputPath = path.resolve(__dirname, "./output", "CONSOLIDATED.xlsx");
    const sheet = workbook.getWorksheet(1);
    const newSheet = workbook.addWorksheet("Consolidated");
    if (!sheet) return;

    const column = sheet.getColumn("D");
    const cell = sheet.getCell('A1');
    const font = cell.font;
    
    const MAX_LINEAS_POR_CELDA = 2;
    const width = column.width || 8.43;
    const fontSize = font?.size || 12;
    const fontName = font?.name || "Arial";

    const approxCharsFit = () => {
        const fontFactors: any = {
            "calibri": 1.00,
            "arial": 1.02,
            "helvetica": 1.02,
            "verdana": 1.10,
            "tahoma": 1.07,
            "times new roman": 0.95,
            "courier new": 0.90,
        };
        const key = fontName.toLowerCase().trim();
        const factor = fontFactors[key] || 1.00;
        const chars = width * (11 / fontSize) * (1 / factor);
        return Math.floor(chars);
    };

    const maxChars = approxCharsFit();

    let TempNameSet: string = "";
    const TempBlockText: FilaTabla[] = [];
    sheet.eachRow((row, rowNumber) => {
        const fila: FilaTabla = {
            Name: row.getCell("C").value?.toString().trim() || "",
            Text: row.getCell("D").value?.toString().trim()
                .replaceAll("\n", " ")
                .replaceAll(".¿", ". ¿")
                .replace(/\.([A-Za-z])/g, '. $1')
                .replace(/ +/g, ' ') || "",
            Row: row  // ✅ Agregado aquí
        };
        if (fila.Text.includes("[")) return;
        const strartTime = row.getCell("A").value?.toString().trim() || "";
        if (strartTime === "Start Time") return;
        if (fila.Name.length === 0) return;
        if (TempNameSet !== fila.Name) {
            if (TempNameSet !== "") {
                const reactionFlasg = TempBlockText.filter(e => e.Text === "REACTION").map(e=>e.Row)
                const cleanBlockText = TempBlockText.filter(e => e.Text !== "REACTION")
                if (cleanBlockText.length !== 0) blockAnalysis(cleanBlockText, reactionFlasg);
                TempBlockText.length = 0;
            }
            TempNameSet = fila.Name;
        }
        fila.Name = TempNameSet;
        TempBlockText.push(fila);
    });

    function blockAnalysis(block: FilaTabla[], reaction : Row[]) {
        const name = block[0].Name;
        const fullText = block.map(fila => fila.Text).join(" - ");

        console.log("_____________________________________________________");
        console.log("Name:" + name);
        // console.log(fullText);

        const result1 = capa1(fullText);
        const result2 = capa2(result1, block);
        const result3 = capa3(result2, maxChars);
        capa4(result3)
    }

    function capa1(fullText: string) {
        const separeMind = fullText.replaceAll(" - ", " ").split("(");

        type BoxDialogType = Array<Partial<Record<"mind" | "normal", string>>>;
        const BoxDialog: BoxDialogType = [];

        if (separeMind.length > 0 && separeMind[0].trim().length > 0) {
            BoxDialog.push({ normal: separeMind[0].trim() });
        }

        for (let i = 1; i < separeMind.length; i++) {
            const element = separeMind[i];

            if (!element.includes(")")) {
                BoxDialog.push({ normal: element.trim() });
                continue;
            }

            const parts = element.trim().split(")");
            BoxDialog.push({ mind: parts[0].trim() });
            if (parts.length > 1 && parts[1].trim().length > 0) {
                BoxDialog.push({ normal: parts[1].trim() });
            }
        }

        const result: ResultType = [];
        const tempAccumulator: string[] = [];
        let flagKey: string = "";

        for (const element of BoxDialog) {
            const key = Object.keys(element)[0] as "mind" | "normal";
            const value = element[key] || "";

            if (flagKey !== key) {
                if (tempAccumulator.length > 0) {
                    result.push({ [flagKey]: tempAccumulator.join(" ") });
                    tempAccumulator.length = 0;
                }
                flagKey = key;
            }

            tempAccumulator.push(value);
        }

        if (tempAccumulator.length > 0) {
            result.push({ [flagKey]: tempAccumulator.join(" ") });
        }

        return result;
    }

    function capa2(boxResult: ResultType, blocks: FilaTabla[]): Capa2Result {
        const blokFrase: Capa2Result = [];

        for (const element of boxResult) {
            const text = Object.values(element)[0] || "";
            const itemScan: ItemScan = {
                text: text,
                type: Object.keys(element)[0] as "mind" | "normal",
                frases: []
            };

            const fraseAsignado: FraseAsignada[] = [];
            const partition = text.split('.');

            for (const frase of partition) {
                const itemFrase: FraseAsignada = {
                    bloque: null,
                    bloqueIndex: null,
                    texto: frase.trim()
                };

                if (itemFrase.texto.length === 0) continue;

                for (const block of blocks) {
                    const btxt = block.Text.replaceAll("(", "").replaceAll(")", "").replaceAll(" - ", " ");
                    if (itemFrase.bloque === null && btxt.includes(itemFrase.texto)) {
                        itemFrase.bloque = block;
                        itemFrase.bloqueIndex = blocks.indexOf(block);
                        break;
                    }
                    if (itemFrase.bloque === null && itemFrase.texto.includes(btxt)) {
                        itemFrase.bloque = block;
                        itemFrase.bloqueIndex = blocks.indexOf(block);
                        break;
                    }

                }
                fraseAsignado.push(itemFrase);
            }

            itemScan.frases = fraseAsignado;
            blokFrase.push(itemScan);
        }

        return blokFrase;
    }

    function capa3(blockText: Capa2Result, maxChars: number) {
        // console.log(blockText);
        const retorno : CeldaConsolidada[] = []
        const packCeldas: FraseAsignada[][] = []
        const totalMaxChars = maxChars * MAX_LINEAS_POR_CELDA
        for (const element of blockText) {
            const celdas = consiliador(element.frases, totalMaxChars, maxChars, element.type)
            packCeldas.push(...celdas)
        }

        for (const element of packCeldas) {
            retorno.push({
                row: element[0].bloque?.Row,
                text: separeToLines(viewPill(element), maxChars)
            })
        }

        function consiliador(
            frases: FraseAsignada[],
            totalMaxChars: number,
            maxLine: number,
            typeItem : "mind" | "normal"
        ): FraseAsignada[][] {

            let preProcesado: FraseAsignada[] = []
            for (const element of frases) {
                const bloques = makeBlocks(element, totalMaxChars, maxLine)
                preProcesado.push(...(bloques.length > 0 ? bloques : [element]))
            }

            const packCeldas: FraseAsignada[][] = []
            let indice = 0;
            while (indice < preProcesado.length) {
                const celda: FraseAsignada[] = [preProcesado[indice]]
                indice++;

                // ✅ Modo codicioso: intenta agregar todas las frases posibles
                while (indice < preProcesado.length) {
                    const prueba = [...celda, preProcesado[indice]]
                    const textoUnido = viewPill(prueba)
                    const lineasResultantes = contarLineas(textoUnido, maxLine)

                    // Verificar límite de líneas
                    if (lineasResultantes > MAX_LINEAS_POR_CELDA) break

                    // ✅ También verificar que no exceda caracteres totales
                    if (textoUnido.length > totalMaxChars) break

                    celda.push(preProcesado[indice])
                    indice++;
                }

                if (celda.length > 0) {
                    celda[0].texto = (typeItem=="mind" ? "(" : "") + celda[0].texto
                    celda[celda.length-1].texto = celda[celda.length-1].texto + (typeItem=="mind" ? ")" : "")
                    packCeldas.push(celda);
                    // const unionTXT = viewPill(celda);
                    // const numLineas = contarLineas(unionTXT, maxLine);
                    // console.log(`\n✅ Celda (${numLineas} líneas, ${unionTXT.length} chars):`);
                    // console.log(separeToLines(unionTXT, maxLine) + "\n");
                }
            }

            return packCeldas;
        }

        function contarLineas(texto: string, maxCharsPerLine: number): number {
            const lineas = separeToLines(texto, maxCharsPerLine);
            return lineas.split("\n").filter(l => l.trim()).length;
        }

        function makeBlocks(frace: FraseAsignada, totalMaxChars: number, maxLine: number) {
            const returnBlocks: FraseAsignada[] = [];

            // Si la frase ya cabe, devuélvela la frace normal
            if (frace.texto.length <= totalMaxChars) return [frace];

            let reciduo = separadorInline(frace, totalMaxChars, maxLine, (section) => {
                returnBlocks.push(section)
            })

            // Protección contra bucle infinito
            let iteraciones = 0;
            const MAX_ITERACIONES = 100;

            while (reciduo.texto.length > totalMaxChars && iteraciones < MAX_ITERACIONES) {
                reciduo = separadorInline(reciduo, totalMaxChars, maxLine, (section) => {
                    returnBlocks.push(section)
                })
                iteraciones++;
            }

            // Agrega el último residuo
            if (reciduo.texto.trim().length > 0) {
                returnBlocks.push(reciduo)
            }

            return returnBlocks;
        }

        function separadorInline(
            frase: FraseAsignada,
            maxChar: number,
            maxLine: number,
            segment: (e: FraseAsignada) => void
        ) {
            const texto = frase.texto;

            if (texto.length <= maxChar) {
                segment({ ...frase });
                return { ...frase, texto: "" };
            }

            // ✅ Estrategia: Buscar el bloque completo más cercano
            const bloqueCompleto = encontrarProximoBloqueCompleto(texto);

            if (bloqueCompleto && bloqueCompleto.fin <= maxChar) {
                // Si hay un bloque completo que cabe en el límite, úsalo
                const recorte = texto.slice(0, bloqueCompleto.fin).trim();
                segment({ ...frase, texto: recorte });
                return { ...frase, texto: texto.slice(bloqueCompleto.fin).trim() };
            }

            // Si no, buscar otros puntos de corte (puntos, comas) ANTES del bloque
            const ventana = texto.slice(0, maxChar);
            const puntosFinal = [...ventana.matchAll(/\.\s/g)].map(m => m.index! + 1);
            const comas = [...ventana.matchAll(/,\s/g)].map(m => m.index! + 1);

            // Si hay un bloque que empieza dentro de la ventana, cortar ANTES
            if (bloqueCompleto && bloqueCompleto.inicio < maxChar) {
                const puntosAntes = [...puntosFinal, ...comas].filter(p => p < bloqueCompleto.inicio);

                if (puntosAntes.length > 0) {
                    const puntoCorte = Math.max(...puntosAntes);
                    const recorte = texto.slice(0, puntoCorte).trim();
                    segment({ ...frase, texto: recorte });
                    return { ...frase, texto: texto.slice(puntoCorte).trim() };
                }
            }

            // Último recurso: cortar por cualquier punto disponible
            const todosPuntos = [...puntosFinal, ...comas];
            const puntoCorte = todosPuntos.length > 0
                ? Math.max(...todosPuntos)
                : ventana.lastIndexOf(' ');

            const recorte = texto.slice(0, puntoCorte > 0 ? puntoCorte : maxChar).trim();
            segment({ ...frase, texto: recorte });
            return { ...frase, texto: texto.slice(recorte.length).trim() };
        }

        // ✅ Encuentra el próximo bloque completo de pregunta/exclamación
        function encontrarProximoBloqueCompleto(texto: string): { inicio: number, fin: number } | null {
            // Buscar ¿...?
            const matchPregunta = texto.match(/¿[^?]*\?/);
            if (matchPregunta && matchPregunta.index !== undefined) {
                return {
                    inicio: matchPregunta.index,
                    fin: matchPregunta.index + matchPregunta[0].length
                };
            }

            // Buscar ¡...!
            const matchExclamacion = texto.match(/¡[^!]*!/);
            if (matchExclamacion && matchExclamacion.index !== undefined) {
                return {
                    inicio: matchExclamacion.index,
                    fin: matchExclamacion.index + matchExclamacion[0].length
                };
            }

            return null;
        }

        function separadorInteligente(texto: string, maxChar: number) {
            const slice = texto.slice(0, maxChar);

            // Jerarquía de puntos de corte (de mejor a peor)
            const prioridades = [
                { regex: /\.\s/g, peso: 100, offset: 1 },      // Punto final
                { regex: /\?\s/g, peso: 90, offset: 1 },       // Pregunta
                { regex: /!\s/g, peso: 90, offset: 1 },        // Exclamación
                { regex: /;\s/g, peso: 80, offset: 1 },        // Punto y coma
                { regex: /,\s/g, peso: 70, offset: 1 },        // Coma
                { regex: /:\s/g, peso: 60, offset: 1 },        // Dos puntos
                { regex: /\s-\s/g, peso: 50, offset: 2 },      // Guion con espacios
                { regex: /\s/g, peso: 10, offset: 0 }          // Último recurso: espacio
            ];

            let mejorCorte = { posicion: -1, peso: -1 };

            for (const prioridad of prioridades) {
                const matches = [...slice.matchAll(prioridad.regex)];
                if (matches.length > 0) {
                    const ultimoMatch = matches[matches.length - 1];
                    const posicion = ultimoMatch.index! + prioridad.offset;

                    if (prioridad.peso > mejorCorte.peso) {
                        mejorCorte = { posicion, peso: prioridad.peso };
                    }
                }
            }

            return mejorCorte.posicion > 0 ? mejorCorte.posicion : maxChar;
        }

        function viewPill(lista: FraseAsignada[]) {
            return (lista.map(e => e.texto).join(". ") + ".").replaceAll("..", ".")
        }

        function separeToLines(texto: string, maxChars: number) {
            let resultado = ""
            let residuo = separador(texto, maxChars, (line: string) => {
                resultado += line + " \n"
            })
            while (residuo.length > 0) {
                residuo = separador(residuo, maxChars, (line: string) => {
                    resultado += line + " \n"
                })
            }
            return (resultado.trim() + ".")
                .replaceAll("..", ".")
                .replaceAll(",.", ",")
                .replaceAll("?.", "?")
                .replaceAll("!.", "!")
        }

        function separador(texto: string, maxChars: number, add: (line: string) => void) {
            if (texto.length > maxChars) {
                const separado = texto.slice(0, maxChars).split(" ");
                const original = texto.split(" ");
                if (original[separado.length - 1].length > 1) separado.pop();
                const line = separado.join(" ");
                add(line.trim());
                return texto.slice(line.length, texto.length).trim() || "";
            } else {
                add(texto.trim())
            }
            return ""
        }

        return retorno;

    }

    function capa4(celdas : CeldaConsolidada[]){

    }

    // await workbook.xlsx.writeFile(outputPath);
}

main().catch(console.error);
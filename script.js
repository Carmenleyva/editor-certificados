/* ======================================================
   ELEMENTOS DE LA BASE DE DATOS
====================================================== */

const subirExcel =
  document.getElementById("subirExcel");

const estadoExcel =
  document.getElementById("estadoExcel");

const selectorParticipante =
  document.getElementById("selectorParticipante");

const codigoActual =
  document.getElementById("codigoActual");

const certificadoAnterior =
  document.getElementById(
    "certificadoAnterior"
  );

const certificadoSiguiente =
  document.getElementById(
    "certificadoSiguiente"
  );

let participantes = [];
let participanteSeleccionado = null;


/* ======================================================
   ELEMENTOS DEL NOMBRE
====================================================== */

const campoNombre =
  document.getElementById("nombre");

const boton =
  document.getElementById("boton");

const resultado =
  document.getElementById("resultado");

const tamanoTexto =
  document.getElementById("tamanoTexto");

const tamanoTextoNumero =
  document.getElementById("tamanoTextoNumero");

const colorTexto =
  document.getElementById("colorTexto");

const fuenteTexto =
  document.getElementById("fuenteTexto");


/* ======================================================
   ELEMENTOS DE LA PLANTILLA
====================================================== */

const subirPlantilla =
  document.getElementById("subirPlantilla");

const plantilla =
  document.getElementById("plantilla");

const certificado =
  document.querySelector(".certificado");


/* ======================================================
   PRIMER PÁRRAFO
====================================================== */

const textoContenido =
  document.getElementById("textoContenido");

const contenidoCertificado =
  document.getElementById("contenidoCertificado");

const alineacionContenido =
  document.getElementById("alineacionContenido");

const tamanoContenido =
  document.getElementById("tamanoContenido");

const tamanoContenidoNumero =
  document.getElementById("tamanoContenidoNumero");

const colorContenido =
  document.getElementById("colorContenido");

const fuenteContenido =
  document.getElementById("fuenteContenido");


/* ======================================================
   SEGUNDO PÁRRAFO
====================================================== */

const textoContenidoDos =
  document.getElementById("textoContenidoDos");

const contenidoCertificadoDos =
  document.getElementById("contenidoCertificadoDos");

const alineacionContenidoDos =
  document.getElementById("alineacionContenidoDos");

const tamanoContenidoDos =
  document.getElementById("tamanoContenidoDos");

const tamanoContenidoDosNumero =
  document.getElementById("tamanoContenidoDosNumero");

const colorContenidoDos =
  document.getElementById("colorContenidoDos");

const fuenteContenidoDos =
  document.getElementById("fuenteContenidoDos");


/* ======================================================
   FECHA
====================================================== */

const textoFecha =
  document.getElementById("textoFecha");

const fechaCertificado =
  document.getElementById("fechaCertificado");

const alineacionFecha =
  document.getElementById("alineacionFecha");

const tamanoFecha =
  document.getElementById("tamanoFecha");

const tamanoFechaNumero =
  document.getElementById("tamanoFechaNumero");

const colorFecha =
  document.getElementById("colorFecha");

const fuenteFecha =
  document.getElementById("fuenteFecha");


/* ======================================================
   OTROS ELEMENTOS
====================================================== */

const botonDescargar =
  document.getElementById("descargar");

const botonDescargarTodos =
  document.getElementById("descargarTodos");

const estadoDescarga =
  document.getElementById("estadoDescarga");

const botonesFormato =
  document.querySelectorAll(".boton-formato");


/* ======================================================
   FUNCIONES AUXILIARES
====================================================== */

function convertirATexto(valor) {
  if (
    valor === undefined ||
    valor === null
  ) {
    return "";
  }

  return String(valor).trim();
}


function normalizarEncabezado(texto) {
  return convertirATexto(texto)
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "_");
}


function obtenerValorFila(fila, nombreColumna) {
  const claveBuscada =
    normalizarEncabezado(nombreColumna);

  const claveEncontrada =
    Object.keys(fila).find(function (clave) {
      return normalizarEncabezado(clave) === claveBuscada;
    });

  if (!claveEncontrada) {
    return "";
  }

  return fila[claveEncontrada];
}


function formatearCodigo(codigo) {
  const texto =
    convertirATexto(codigo);

  if (texto === "") {
    return "";
  }

  if (/^\d+$/.test(texto)) {
    return texto.padStart(3, "0");
  }

  return texto;
}


function limpiarNombreArchivo(texto) {
  return convertirATexto(texto)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\p{L}\p{N}\s-]/gu, "")
    .trim()
    .replace(/\s+/g, "-");
}


function colocarTextoEnEditor(
  editor,
  texto
) {
  editor.textContent =
    convertirATexto(texto);
}


/* ======================================================
   LEER EL EXCEL
====================================================== */

subirExcel.addEventListener(
  "change",
  function () {
    const archivo =
      subirExcel.files[0];

    if (!archivo) {
      return;
    }

    estadoExcel.textContent =
      "Leyendo el archivo...";

    const lector =
      new FileReader();

    lector.onload =
      function (evento) {
        try {
          const datos =
            new Uint8Array(
              evento.target.result
            );

          const libro =
            XLSX.read(datos, {
              type: "array"
            });

          const nombrePrimeraHoja =
            libro.SheetNames[0];

          const hoja =
            libro.Sheets[
              nombrePrimeraHoja
            ];

          const filas =
            XLSX.utils.sheet_to_json(
              hoja,
              {
                defval: "",
                raw: false
              }
            );

          participantes =
            filas
              .map(function (fila, indice) {
                return {
                  codigo:
                    obtenerValorFila(
                      fila,
                      "codigo"
                    ),

                  nombresApellidos:
                    obtenerValorFila(
                      fila,
                      "nombres_apellidos"
                    ),

                  texto1:
                    obtenerValorFila(
                      fila,
                      "texto_1"
                    ),

                  texto2:
                    obtenerValorFila(
                      fila,
                      "texto_2"
                    ),

                  fecha:
                    obtenerValorFila(
                      fila,
                      "fecha"
                    ),

                  ordenOriginal:
                    indice
                };
              })
              .filter(function (persona) {
                return (
                  convertirATexto(
                    persona.nombresApellidos
                  ) !== ""
                );
              });

          participantes.sort(
            function (a, b) {
              const codigoA =
                Number(a.codigo);

              const codigoB =
                Number(b.codigo);

              const ambosNumericos =
                Number.isFinite(codigoA) &&
                Number.isFinite(codigoB);

              if (ambosNumericos) {
                return codigoA - codigoB;
              }

              return convertirATexto(
                a.codigo
              ).localeCompare(
                convertirATexto(
                  b.codigo
                ),
                "es",
                {
                  numeric: true
                }
              );
            }
          );

          if (
            participantes.length === 0
          ) {
            estadoExcel.textContent =
              "El archivo no contiene participantes válidos.";

            selectorParticipante.disabled =
              true;

            return;
          }

          llenarSelectorParticipantes();

          estadoExcel.textContent =
            participantes.length +
            " participantes cargados correctamente.";

          selectorParticipante.disabled =
  false;

botonDescargarTodos.disabled =
  false;

selectorParticipante.value =
  "0";

          cargarParticipante(0);
        } catch (error) {
          console.error(error);

          participantes = [];

          selectorParticipante.innerHTML =
            '<option value="">No se pudo leer el archivo</option>';

          selectorParticipante.disabled =
            true;

          estadoExcel.textContent =
            "No se pudo leer el Excel. Revisa los encabezados.";
        }
      };

    lector.onerror =
      function () {
        estadoExcel.textContent =
          "Ocurrió un error al abrir el archivo.";
      };

    lector.readAsArrayBuffer(
      archivo
    );
  }
);


/* ======================================================
   LLENAR EL SELECTOR DE PARTICIPANTES
====================================================== */

function llenarSelectorParticipantes() {
  selectorParticipante.innerHTML = "";

  participantes.forEach(
    function (persona, indice) {
      const opcion =
        document.createElement(
          "option"
        );

      const codigo =
        formatearCodigo(
          persona.codigo
        );

      opcion.value =
        indice;

      opcion.textContent =
        codigo +
        " - " +
        convertirATexto(
          persona.nombresApellidos
        );

      selectorParticipante.appendChild(
        opcion
      );
    }
  );
}


/* ======================================================
   CARGAR UN PARTICIPANTE DEL EXCEL
====================================================== */

function cargarParticipante(indice) {
  const persona =
    participantes[indice];

  if (!persona) {
    return;
  }

  participanteSeleccionado =
    persona;

  const codigo =
    formatearCodigo(
      persona.codigo
    );

  codigoActual.textContent =
    codigo || "—";

  campoNombre.value =
    convertirATexto(
      persona.nombresApellidos
    );

  resultado.textContent =
    convertirATexto(
      persona.nombresApellidos
    ) ||
    "Aquí aparecerá el nombre";

  colocarTextoEnEditor(
    textoContenido,
    persona.texto1
  );

  colocarTextoEnEditor(
    textoContenidoDos,
    persona.texto2
  );

  colocarTextoEnEditor(
    textoFecha,
    persona.fecha
  );

  actualizarPrimerParrafo();
  actualizarSegundoParrafo();
  actualizarFecha();
  actualizarBotonesNavegacion();
}


function actualizarBotonesNavegacion() {
  const indiceActual =
    Number(
      selectorParticipante.value
    );

  certificadoAnterior.disabled =
    participantes.length === 0 ||
    indiceActual <= 0;

  certificadoSiguiente.disabled =
    participantes.length === 0 ||
    indiceActual >=
      participantes.length - 1;
}


selectorParticipante.addEventListener(
  "change",
  function () {
    const indice =
      Number(
        selectorParticipante.value
      );

    cargarParticipante(indice);
  }
);


certificadoAnterior.addEventListener(
  "click",
  function () {
    const indiceActual =
      Number(
        selectorParticipante.value
      );

    if (indiceActual <= 0) {
      return;
    }

    const nuevoIndice =
      indiceActual - 1;

    selectorParticipante.value =
      nuevoIndice;

    cargarParticipante(
      nuevoIndice
    );
  }
);


certificadoSiguiente.addEventListener(
  "click",
  function () {
    const indiceActual =
      Number(
        selectorParticipante.value
      );

    if (
      indiceActual >=
      participantes.length - 1
    ) {
      return;
    }

    const nuevoIndice =
      indiceActual + 1;

    selectorParticipante.value =
      nuevoIndice;

    cargarParticipante(
      nuevoIndice
    );
  }
);


/* ======================================================
   MOSTRAR EL NOMBRE MANUALMENTE
====================================================== */

function mostrarNombre() {
  const nombreEscrito =
    campoNombre.value.trim();

  if (nombreEscrito === "") {
    resultado.textContent =
      "Escribe un nombre";

    return;
  }

  resultado.textContent =
    nombreEscrito;
}


boton.addEventListener(
  "click",
  mostrarNombre
);


campoNombre.addEventListener(
  "input",
  function () {
    resultado.textContent =
      campoNombre.value.trim() ||
      "Aquí aparecerá el nombre";
  }
);


/* ======================================================
   CARGAR LA PLANTILLA
====================================================== */

subirPlantilla.addEventListener(
  "change",
  function () {
    const archivo =
      subirPlantilla.files[0];

    if (!archivo) {
      return;
    }

    plantilla.src =
      URL.createObjectURL(
        archivo
      );
  }
);


/* ======================================================
   TAMAÑO DEL NOMBRE
====================================================== */

function actualizarTamanoNombre(valor) {
  const numero =
    Number(valor);

  if (!Number.isFinite(numero)) {
    return;
  }

  const tamano =
    Math.min(
      60,
      Math.max(2, numero)
    );

  resultado.style.fontSize =
    tamano + "px";

  tamanoTexto.value =
    tamano;

  tamanoTextoNumero.value =
    tamano;
}


tamanoTexto.addEventListener(
  "input",
  function () {
    actualizarTamanoNombre(
      tamanoTexto.value
    );
  }
);


tamanoTextoNumero.addEventListener(
  "input",
  function () {
    actualizarTamanoNombre(
      tamanoTextoNumero.value
    );
  }
);


/* ======================================================
   COLOR Y FUENTE DEL NOMBRE
====================================================== */

colorTexto.addEventListener(
  "input",
  function () {
    resultado.style.color =
      colorTexto.value;
  }
);


fuenteTexto.addEventListener(
  "change",
  function () {
    resultado.style.fontFamily =
      fuenteTexto.value;
  }
);


/* ======================================================
   SINCRONIZAR LOS EDITORES
====================================================== */

function actualizarPrimerParrafo() {
  contenidoCertificado.innerHTML =
    textoContenido.innerHTML;
}


function actualizarSegundoParrafo() {
  contenidoCertificadoDos.innerHTML =
    textoContenidoDos.innerHTML;
}


function actualizarFecha() {
  fechaCertificado.innerHTML =
    textoFecha.innerHTML;
}


textoContenido.addEventListener(
  "input",
  actualizarPrimerParrafo
);


textoContenidoDos.addEventListener(
  "input",
  actualizarSegundoParrafo
);


textoFecha.addEventListener(
  "input",
  actualizarFecha
);


/* ======================================================
   PEGAR TEXTO LIMPIO
====================================================== */

function activarPegadoLimpio(editor) {
  editor.addEventListener(
    "paste",
    function (evento) {
      evento.preventDefault();

      const textoPlano =
        evento.clipboardData.getData(
          "text/plain"
        );

      document.execCommand(
        "insertText",
        false,
        textoPlano
      );
    }
  );
}


activarPegadoLimpio(
  textoContenido
);

activarPegadoLimpio(
  textoContenidoDos
);

activarPegadoLimpio(
  textoFecha
);


/* ======================================================
   BOTONES DE NEGRITA, CURSIVA Y SUBRAYADO
====================================================== */

botonesFormato.forEach(
  function (botonFormato) {
    botonFormato.addEventListener(
      "mousedown",
      function (evento) {
        evento.preventDefault();
      }
    );

    botonFormato.addEventListener(
      "click",
      function () {
        const idEditor =
          botonFormato.dataset.editor;

        const comando =
          botonFormato.dataset.comando;

        const editor =
          document.getElementById(
            idEditor
          );

        if (!editor) {
          return;
        }

        editor.focus();

        document.execCommand(
          comando,
          false,
          null
        );

        if (
          idEditor ===
          "textoContenido"
        ) {
          actualizarPrimerParrafo();
        }

        if (
          idEditor ===
          "textoContenidoDos"
        ) {
          actualizarSegundoParrafo();
        }

        if (
          idEditor ===
          "textoFecha"
        ) {
          actualizarFecha();
        }
      }
    );
  }
);


/* ======================================================
   PRIMER PÁRRAFO
====================================================== */

alineacionContenido.addEventListener(
  "change",
  function () {
    const alineacion =
      alineacionContenido.value;

    textoContenido.style.textAlign =
      alineacion;

    contenidoCertificado.style.textAlign =
      alineacion;
  }
);


function actualizarTamanoContenido(valor) {
  const numero =
    Number(valor);

  if (!Number.isFinite(numero)) {
    return;
  }

  const tamano =
    Math.min(
      40,
      Math.max(2, numero)
    );

  textoContenido.style.fontSize =
    tamano + "px";

  contenidoCertificado.style.fontSize =
    tamano + "px";

  tamanoContenido.value =
    tamano;

  tamanoContenidoNumero.value =
    tamano;
}


tamanoContenido.addEventListener(
  "input",
  function () {
    actualizarTamanoContenido(
      tamanoContenido.value
    );
  }
);


tamanoContenidoNumero.addEventListener(
  "input",
  function () {
    actualizarTamanoContenido(
      tamanoContenidoNumero.value
    );
  }
);


colorContenido.addEventListener(
  "input",
  function () {
    const color =
      colorContenido.value;

    textoContenido.style.color =
      color;

    contenidoCertificado.style.color =
      color;
  }
);


fuenteContenido.addEventListener(
  "change",
  function () {
    const fuente =
      fuenteContenido.value;

    textoContenido.style.fontFamily =
      fuente;

    contenidoCertificado.style.fontFamily =
      fuente;
  }
);


/* ======================================================
   SEGUNDO PÁRRAFO
====================================================== */

alineacionContenidoDos.addEventListener(
  "change",
  function () {
    const alineacion =
      alineacionContenidoDos.value;

    textoContenidoDos.style.textAlign =
      alineacion;

    contenidoCertificadoDos.style.textAlign =
      alineacion;
  }
);


function actualizarTamanoContenidoDos(valor) {
  const numero =
    Number(valor);

  if (!Number.isFinite(numero)) {
    return;
  }

  const tamano =
    Math.min(
      40,
      Math.max(2, numero)
    );

  textoContenidoDos.style.fontSize =
    tamano + "px";

  contenidoCertificadoDos.style.fontSize =
    tamano + "px";

  tamanoContenidoDos.value =
    tamano;

  tamanoContenidoDosNumero.value =
    tamano;
}


tamanoContenidoDos.addEventListener(
  "input",
  function () {
    actualizarTamanoContenidoDos(
      tamanoContenidoDos.value
    );
  }
);


tamanoContenidoDosNumero.addEventListener(
  "input",
  function () {
    actualizarTamanoContenidoDos(
      tamanoContenidoDosNumero.value
    );
  }
);


colorContenidoDos.addEventListener(
  "input",
  function () {
    const color =
      colorContenidoDos.value;

    textoContenidoDos.style.color =
      color;

    contenidoCertificadoDos.style.color =
      color;
  }
);


fuenteContenidoDos.addEventListener(
  "change",
  function () {
    const fuente =
      fuenteContenidoDos.value;

    textoContenidoDos.style.fontFamily =
      fuente;

    contenidoCertificadoDos.style.fontFamily =
      fuente;
  }
);


/* ======================================================
   FECHA
====================================================== */

alineacionFecha.addEventListener(
  "change",
  function () {
    const alineacion =
      alineacionFecha.value;

    textoFecha.style.textAlign =
      alineacion;

    fechaCertificado.style.textAlign =
      alineacion;
  }
);


function actualizarTamanoFecha(valor) {
  const numero =
    Number(valor);

  if (!Number.isFinite(numero)) {
    return;
  }

  const tamano =
    Math.min(
      40,
      Math.max(2, numero)
    );

  textoFecha.style.fontSize =
    tamano + "px";

  fechaCertificado.style.fontSize =
    tamano + "px";

  tamanoFecha.value =
    tamano;

  tamanoFechaNumero.value =
    tamano;
}


tamanoFecha.addEventListener(
  "input",
  function () {
    actualizarTamanoFecha(
      tamanoFecha.value
    );
  }
);


tamanoFechaNumero.addEventListener(
  "input",
  function () {
    actualizarTamanoFecha(
      tamanoFechaNumero.value
    );
  }
);


colorFecha.addEventListener(
  "input",
  function () {
    const color =
      colorFecha.value;

    textoFecha.style.color =
      color;

    fechaCertificado.style.color =
      color;
  }
);


fuenteFecha.addEventListener(
  "change",
  function () {
    const fuente =
      fuenteFecha.value;

    textoFecha.style.fontFamily =
      fuente;

    fechaCertificado.style.fontFamily =
      fuente;
  }
);


/* ======================================================
   MOVER LOS TEXTOS DENTRO DEL CERTIFICADO
====================================================== */

function activarArrastre(elemento) {
  let arrastrando =
    false;

  elemento.style.touchAction =
    "none";

  elemento.addEventListener(
    "pointerdown",
    function (evento) {
      evento.preventDefault();
      evento.stopPropagation();

      arrastrando =
        true;

      elemento.style.cursor =
        "grabbing";

      elemento.setPointerCapture(
        evento.pointerId
      );
    }
  );

  elemento.addEventListener(
    "pointermove",
    function (evento) {
      if (!arrastrando) {
        return;
      }

      const rectangulo =
        certificado.getBoundingClientRect();

      let posicionX =
        (
          (
            evento.clientX -
            rectangulo.left
          ) /
          rectangulo.width
        ) * 100;

      let posicionY =
        (
          (
            evento.clientY -
            rectangulo.top
          ) /
          rectangulo.height
        ) * 100;

      posicionX =
        Math.max(
          0,
          Math.min(
            100,
            posicionX
          )
        );

      posicionY =
        Math.max(
          0,
          Math.min(
            100,
            posicionY
          )
        );

      elemento.style.left =
        posicionX + "%";

      elemento.style.top =
        posicionY + "%";
    }
  );

  elemento.addEventListener(
    "pointerup",
    function (evento) {
      arrastrando =
        false;

      elemento.style.cursor =
        "grab";

      if (
        elemento.hasPointerCapture(
          evento.pointerId
        )
      ) {
        elemento.releasePointerCapture(
          evento.pointerId
        );
      }
    }
  );

  elemento.addEventListener(
    "pointercancel",
    function () {
      arrastrando =
        false;

      elemento.style.cursor =
        "grab";
    }
  );
}


activarArrastre(
  resultado
);

activarArrastre(
  contenidoCertificado
);

activarArrastre(
  contenidoCertificadoDos
);

activarArrastre(
  fechaCertificado
);


/* ======================================================
   DESCARGAR CERTIFICADO
====================================================== */

botonDescargar.addEventListener(
  "click",
  function () {
    if (!plantilla.src) {
      alert(
        "Primero selecciona una plantilla."
      );

      return;
    }

    html2canvas(
      certificado,
      {
        scale: 2,
        useCORS: true,
        backgroundColor: null
      }
    )
      .then(function (canvas) {
        const enlace =
          document.createElement(
            "a"
          );

        const nombrePersona =
          campoNombre.value.trim() ||
          "certificado";

        let codigo =
          "";

        if (
          participanteSeleccionado
        ) {
          codigo =
            formatearCodigo(
              participanteSeleccionado.codigo
            );
        }

        const nombreSeguro =
          limpiarNombreArchivo(
            nombrePersona
          );

        if (codigo !== "") {
          enlace.download =
            codigo +
            "-" +
            nombreSeguro +
            ".png";
        } else {
          enlace.download =
            "certificado-" +
            nombreSeguro +
            ".png";
        }

        enlace.href =
          canvas.toDataURL(
            "image/png"
          );

        enlace.click();
      })
      .catch(function (error) {
        console.error(error);

        alert(
          "No se pudo descargar el certificado."
        );
      });
  }
);

/* ======================================================
   DESCARGAR TODOS LOS CERTIFICADOS
====================================================== */

function esperarRenderizado() {
  return new Promise(function (resolver) {
    requestAnimationFrame(function () {
      requestAnimationFrame(resolver);
    });
  });
}

function convertirCanvasABlob(canvas) {
  return new Promise(function (resolver, rechazar) {
    canvas.toBlob(
      function (blob) {
        if (blob) {
          resolver(blob);
        } else {
          rechazar(
            new Error(
              "No se pudo crear la imagen del certificado."
            )
          );
        }
      },
      "image/png"
    );
  });
}

function crearCopiaCertificado(persona) {
  const copia =
    certificado.cloneNode(true);

  copia.style.position =
    "fixed";

  copia.style.left =
    "-100000px";

  copia.style.top =
    "0";

  copia.style.width =
    certificado.offsetWidth + "px";

  copia.style.margin =
    "0";

  const nombreCopia =
    copia.querySelector("#resultado");

  const primerParrafoCopia =
    copia.querySelector(
      "#contenidoCertificado"
    );

  const segundoParrafoCopia =
    copia.querySelector(
      "#contenidoCertificadoDos"
    );

  const fechaCopia =
    copia.querySelector(
      "#fechaCertificado"
    );

  nombreCopia.textContent =
    convertirATexto(
      persona.nombresApellidos
    ) || "Aquí aparecerá el nombre";

  primerParrafoCopia.textContent =
    convertirATexto(
      persona.texto1
    );

  segundoParrafoCopia.textContent =
    convertirATexto(
      persona.texto2
    );

  fechaCopia.textContent =
    convertirATexto(
      persona.fecha
    );

  document.body.appendChild(
    copia
  );

  return copia;
}

botonDescargarTodos.addEventListener(
  "click",
  async function () {
    if (!plantilla.src) {
      alert(
        "Primero selecciona una plantilla."
      );

      return;
    }

    if (participantes.length === 0) {
      alert(
        "Primero carga un archivo Excel con participantes."
      );

      return;
    }

    botonDescargarTodos.disabled =
      true;

    estadoDescarga.textContent =
      "Preparando los certificados...";

    try {
      const zip =
        new JSZip();

      for (
        let indice = 0;
        indice < participantes.length;
        indice += 1
      ) {
        const persona =
          participantes[indice];

        estadoDescarga.textContent =
          "Generando certificado " +
          (indice + 1) +
          " de " +
          participantes.length +
          "...";

        const copia =
          crearCopiaCertificado(
            persona
          );

        await esperarRenderizado();

        const canvas =
          await html2canvas(
            copia,
            {
              scale: 2,
              useCORS: true,
              backgroundColor: null
            }
          );

        copia.remove();

        const blob =
          await convertirCanvasABlob(
            canvas
          );

        const codigo =
          formatearCodigo(
            persona.codigo
          );

        const nombreSeguro =
          limpiarNombreArchivo(
            persona.nombresApellidos
          ) || "certificado";

        const nombreArchivo =
          codigo !== ""
            ? codigo +
              "-" +
              nombreSeguro +
              ".png"
            : "certificado-" +
              nombreSeguro +
              ".png";

        zip.file(
          nombreArchivo,
          blob
        );
      }

      estadoDescarga.textContent =
        "Creando el archivo ZIP...";

      const archivoZip =
        await zip.generateAsync({
          type: "blob"
        });

      const enlace =
        document.createElement("a");

      const urlZip =
        URL.createObjectURL(
          archivoZip
        );

      enlace.href =
        urlZip;

      enlace.download =
        "certificados.zip";

      document.body.appendChild(
        enlace
      );

      enlace.click();
      enlace.remove();

      URL.revokeObjectURL(
        urlZip
      );

      estadoDescarga.textContent =
        participantes.length +
        " certificados descargados correctamente.";
    } catch (error) {
      console.error(error);

      estadoDescarga.textContent =
        "No se pudieron generar todos los certificados.";

      alert(
        "No se pudieron descargar todos los certificados."
      );
    } finally {
      botonDescargarTodos.disabled =
        false;
    }
  }
);

/* ======================================================
   SINCRONIZACIÓN INICIAL
====================================================== */

textoContenido.style.color =
  colorContenido.value;

textoContenidoDos.style.color =
  colorContenidoDos.value;

textoFecha.style.color =
  colorFecha.value;

actualizarPrimerParrafo();
actualizarSegundoParrafo();
actualizarFecha();

actualizarTamanoNombre(
  tamanoTexto.value
);

actualizarTamanoContenido(
  tamanoContenido.value
);

actualizarTamanoContenidoDos(
  tamanoContenidoDos.value
);

actualizarTamanoFecha(
  tamanoFecha.value
);
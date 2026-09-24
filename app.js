let estudiantes = [];


// ==========================================
// CARGAR EXCEL
// ==========================================

async function cargarExcel() {

    try {

        const respuesta = await fetch("estudiantes.xlsx");

        if (!respuesta.ok) {
            throw new Error("No se pudo cargar estudiantes.xlsx");
        }

        const archivo = await respuesta.arrayBuffer();

        const libro = XLSX.read(archivo, {
            type: "array"
        });

        const nombreHoja = libro.SheetNames[0];

        const hoja = libro.Sheets[nombreHoja];

        estudiantes = XLSX.utils.sheet_to_json(hoja, {
            defval: ""
        });

        console.log(
            "Excel cargado correctamente:",
            estudiantes.length,
            "registros"
        );

    } catch (error) {

        console.error("Error cargando Excel:", error);

        document.getElementById("mensajeError").textContent =
            "No se pudo cargar el archivo estudiantes.xlsx.";

        document.getElementById("error").style.display = "flex";
    }
}


// ==========================================
// FORMULARIO
// ==========================================

const formulario =
    document.getElementById("formulario");


formulario.addEventListener("submit", function(event) {

    event.preventDefault();

    const buscar =
        document.getElementById("buscar").value.trim();

    const resultado =
        document.getElementById("resultado");

    const error =
        document.getElementById("error");

    const mensajeError =
        document.getElementById("mensajeError");


    resultado.style.display = "none";
    error.style.display = "none";


    if (buscar === "") {

        mensajeError.textContent =
            "Por favor, ingrese un CI o Registro Universitario.";

        error.style.display = "flex";

        return;
    }


    // ==========================================
    // BUSCAR POR CI O RU
    // ==========================================

    const estudiante = estudiantes.find(function(persona) {

        const ci =
            String(persona.CI || "").trim();

        const ru =
            String(persona.RU || "").trim();

        return ci === buscar || ru === buscar;

    });


    // ==========================================
    // MOSTRAR DATOS
    // ==========================================

    if (estudiante) {

        document.getElementById("numero").textContent =
            estudiante.Nro || "";


        document.getElementById("ci").textContent =
            estudiante.CI || "";


        document.getElementById("ru").textContent =
            estudiante.RU || "";


        let nombreCompleto = "";

        if (estudiante.PATERNO) {
            nombreCompleto += estudiante.PATERNO + " ";
        }

        if (estudiante.MATERNO) {
            nombreCompleto += estudiante.MATERNO + " ";
        }

        if (estudiante.NOMBRE) {
            nombreCompleto += estudiante.NOMBRE;
        }


        document.getElementById("nombres").textContent =
            nombreCompleto.trim();


        document.getElementById("mesa").textContent =
            estudiante.MESA || "";


        resultado.style.display = "block";


    } else {

        mensajeError.textContent =
            "No se encontró ningún estudiante con el CI o Registro Universitario ingresado.";

        error.style.display = "flex";
    }

});


// ==========================================
// BOTÓN LIMPIAR
// ==========================================

function limpiarFormulario() {

    document.getElementById("buscar").value = "";

    document.getElementById("resultado").style.display =
        "none";

    document.getElementById("error").style.display =
        "none";

    document.getElementById("buscar").focus();
}


// ==========================================
// INICIAR
// ==========================================

cargarExcel();
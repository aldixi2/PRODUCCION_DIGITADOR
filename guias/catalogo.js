// catalogos.js
// Listas controladas (código + nombre) para que el personal no escriba “a su gusto”.

window.CATALOGOS = {
  dx018: [
    { code: "Z300", name: "Consejo y asesoramiento general sobre la anticoncepción" },
    { code: "Z309", name: "Asistencia para la anticoncepción, no especificada" },
    { code: "Z304", name: "Supervisión del uso de drogas anticonceptivas (hormonales)" },
    { code: "Z308", name: "Otras atenciones especificadas para la anticoncepción (oral emergencia, implantes)" },
    { code: "Z301", name: "Inserción de dispositivo anticonceptivo (DIU)" },
    { code: "Z3051", name: "Retiro de DIU" },
    { code: "Z305", name: "Supervisión del uso de DIU" }
  ],

  proc018: [
    { code: "99402.04", name: "Orientación/consejería en planificación familiar" },
    { code: "99208",    name: "Atención en planificación familiar y salud reproductiva (cuando se entrega método)" },
    { code: "90782",    name: "Inyección profiláctica/diagnóstica/terapéutica; SC o IM" },
    { code: "11975",    name: "Inserción de cápsulas anticonceptivas implantables" },
    { code: "11977",    name: "Retiro e inserción de cápsulas contraceptivas implantables" },
    { code: "90471",    name: "Administración de inmunización (1ra vacuna)" },
    { code: "58300",    name: "Inserción de dispositivo intrauterino (DIU)" },
    { code: "58301",    name: "Remoción de dispositivo intrauterino (DIU)" }
  ],

  // Mini catálogo demo (luego lo llenamos desde tu Excel grande)
  medsDemo: [
    { code: "04677", name: "Metamizol sódico INY 1 g/2 mL" },
    { code: "05335", name: "Paracetamol TAB 500 mg" },
    { code: "04514", name: "Loratadina TAB 10 mg" }
  ],

  insDemo: [
    { code: "08054", name: "Preservativos sin nonoxinol (x10)" },
    { code: "08068", name: "Dispositivo intrauterino de cobre" }
  ]
};

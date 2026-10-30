// Validación de la normalización de cheques del dashboard (income.js).
//
// Ejecutar desde la raíz del proyecto:
//
//     node tests/cheques_frontend.test.js
//
// Comprueba que la tabla "Movimientos de cheques" muestra los mismos números
// que el backend de Python (tests/test_cheques.py).

const fs = require('fs');
const path = require('path');

const archivo = path.join(__dirname, '..', 'static', 'assets', 'js', 'ingreso', 'income.js');
const fuente = fs.readFileSync(archivo, 'utf8');

// Extrae una función del archivo sin ejecutar el resto del script
// (income.js toca el DOM al cargarse, por eso no se puede requerir completo).
function extraerFuncion(nombre) {
  const inicio = fuente.indexOf(`function ${nombre}(`);
  if (inicio === -1) throw new Error(`No se encontró la función ${nombre}`);

  let llaves = 0;
  let abierto = false;
  for (let i = inicio; i < fuente.length; i++) {
    const caracter = fuente[i];
    if (caracter === '{') { llaves++; abierto = true; }
    if (caracter === '}') {
      llaves--;
      if (abierto && llaves === 0) return fuente.slice(inicio, i + 1);
    }
  }
  throw new Error(`No se pudo extraer la función ${nombre}`);
}

const contexto = {};
for (const nombre of ['normalizeCheckNumber', 'transactionValue', 'chequeText', 'chequeType', 'isChequeTransaction']) {
  // Se compilan en el ámbito global para que chequeType() encuentre a
  // chequeText()/transactionValue() igual que en el navegador.
  // eslint-disable-next-line no-new-func
  contexto[nombre] = new Function(`${extraerFuncion(nombre)}; return ${nombre};`)();
  globalThis[nombre] = contexto[nombre];
}

const { normalizeCheckNumber, chequeType, isChequeTransaction } = contexto;

let fallos = 0;
function igual(real, esperado, caso) {
  const ok = real === esperado;
  if (!ok) fallos++;
  console.log(`${ok ? '  OK ' : 'FAIL '} ${caso}: ${JSON.stringify(real)}${ok ? '' : ` (esperado ${JSON.stringify(esperado)})`}`);
}

console.log('normalizeCheckNumber');
igual(normalizeCheckNumber(826486.0), '826486', 'celda numérica 826486.0');
igual(normalizeCheckNumber('826486.0'), '826486', 'texto "826486.0"');
igual(normalizeCheckNumber('000011666'), '11666', 'ceros de relleno');
igual(normalizeCheckNumber('N° 12345'), '12345', 'símbolo N°');
igual(normalizeCheckNumber('Nº12345'), '12345', 'símbolo Nº');
igual(normalizeCheckNumber('No. 1234'), '1234', 'etiqueta No.');
igual(normalizeCheckNumber('#1234'), '1234', 'símbolo #');
igual(normalizeCheckNumber('CHQ No.1234'), '1234', 'abreviatura CHQ');
igual(normalizeCheckNumber('CHQ#1234'), '1234', 'CHQ pegado al número');
igual(normalizeCheckNumber('1234-A'), '1234-A', 'sufijo alfanumérico');
igual(normalizeCheckNumber('4321/AB'), '4321/AB', 'número con barra');
igual(normalizeCheckNumber('12/03/2026'), '', 'fecha descartada');
igual(normalizeCheckNumber('Pago'), '', 'sin número');
igual(normalizeCheckNumber(''), '', 'vacío');
igual(normalizeCheckNumber(null), '', 'null');

console.log('\nisChequeTransaction');
igual(isChequeTransaction({ isCheck: true }), true, 'isCheck del backend');
igual(isChequeTransaction({ checkNumber: '11666' }), true, 'con checkNumber');
igual(isChequeTransaction({ description: 'PAGO CHEQUE No. 000011666' }), true, 'CHEQUE completo');
igual(isChequeTransaction({ description: 'CHQ#12345 PAGO PROVEEDOR' }), true, 'CHQ con símbolo');
igual(isChequeTransaction({ description: 'CK12345' }), true, 'CK pegado');
igual(isChequeTransaction({ description: 'CHKS 9988' }), true, 'CHKS');
igual(isChequeTransaction({ description: 'N° CHEQUE 3344' }), true, 'N° antes de CHEQUE');
igual(isChequeTransaction({ description: 'TRANSFERENCIA A PROVEEDOR' }), false, 'transferencia');
igual(isChequeTransaction({ description: 'DEPOSITO CLIENTE' }), false, 'depósito');
// "GRAVABLOCK" contiene "CK": no debe detectarse como cheque.
igual(isChequeTransaction({ description: 'T365 DE: INDUSTRIAS GRAVABLOCK', reference: '912406873' }), false, 'falso positivo GRAVABLOCK');

console.log('\nchequeType');
igual(chequeType({ checkType: 'EGRESO', debit: 100 }).type, 'EGRESO', 'declarado EGRESO');
igual(chequeType({ checkType: 'INGRESO', credit: 100 }).type, 'INGRESO', 'declarado INGRESO');
igual(chequeType({ description: 'RECHAZO CHEQUE 4455', checkType: 'EGRESO' }).type, 'DEVUELTO', 'rechazo');
igual(chequeType({ description: 'DEV CHEQ 4455', checkType: 'EGRESO' }).type, 'DEVUELTO', 'abreviatura DEV');
igual(chequeType({ description: 'REVERSADO CHEQUE 1234' }).type, 'DEVUELTO', 'reversado');
igual(chequeType({ debit: 100 }).type, 'EGRESO', 'respaldo por débito');
igual(chequeType({ credit: 100 }).type, 'INGRESO', 'respaldo por crédito');

console.log(`\n${fallos ? `${fallos} falla(s)` : 'Todo cuadra'}`);
process.exit(fallos ? 1 : 0);

#!/usr/bin/env node

/**
 * Script completo para probar han-excel-builder
 * Ejecuta todos los tests disponibles
 */

const { execSync } = require('child_process');
const fs = require('fs');
const path = require('path');

console.log('🧪 ========================================');
console.log('🧪 TEST COMPLETO DE HAN-EXCEL-BUILDER');
console.log('🧪 ========================================\n');

let allTestsPassed = true;
const testResults = [];

// Función para ejecutar un test
function runTest(name, command) {
    console.log(`📋 Ejecutando: ${name}`);
    console.log(`🔧 Comando: ${command}\n`);
    
    try {
        const startTime = Date.now();
        execSync(command, { stdio: 'inherit' });
        const endTime = Date.now();
        const duration = endTime - startTime;
        
        console.log(`✅ ${name} - PASÓ (${duration}ms)\n`);
        testResults.push({ name, status: 'PASSED', duration });
        return true;
    } catch (error) {
        console.log(`❌ ${name} - FALLÓ\n`);
        testResults.push({ name, status: 'FAILED', error: error.message });
        allTestsPassed = false;
        return false;
    }
}

// Función para verificar archivos
function checkFile(filename, description) {
    console.log(`📁 Verificando: ${description}`);
    
    if (fs.existsSync(filename)) {
        const stats = fs.statSync(filename);
        const sizeKB = (stats.size / 1024).toFixed(2);
        console.log(`✅ ${filename} existe (${sizeKB} KB)\n`);
        testResults.push({ name: description, status: 'PASSED', fileSize: sizeKB });
        return true;
    } else {
        console.log(`❌ ${filename} no existe\n`);
        testResults.push({ name: description, status: 'FAILED', error: 'Archivo no encontrado' });
        allTestsPassed = false;
        return false;
    }
}

// Función para verificar build
function checkBuild() {
    console.log('🔨 Verificando build...');
    
    const distFiles = [
        'dist/han-excel.es.js',
        'dist/han-excel.cjs.js',
        'dist/index.d.ts'
    ];
    
    let buildOk = true;
    distFiles.forEach(file => {
        if (!fs.existsSync(file)) {
            console.log(`❌ ${file} no existe`);
            buildOk = false;
        }
    });
    
    if (buildOk) {
        console.log('✅ Build completado correctamente\n');
        testResults.push({ name: 'Build', status: 'PASSED' });
    } else {
        console.log('❌ Build incompleto\n');
        testResults.push({ name: 'Build', status: 'FAILED' });
        allTestsPassed = false;
    }
    
    return buildOk;
}

// Ejecutar tests
console.log('🚀 INICIANDO TESTS...\n');

// 1. Verificar build
checkBuild();

// 2. Test básico
runTest('Test Básico', 'npx tsx test-simple.ts');

// 3. Test completo
runTest('Test Completo', 'npx tsx test-complete.ts');

// 4. Verificar archivos generados
checkFile('test-report-complete.xlsx', 'Archivo Excel generado');

// 5. Verificar que el archivo es válido
if (fs.existsSync('test-report-complete.xlsx')) {
    const buffer = fs.readFileSync('test-report-complete.xlsx');
    const isValidExcel = buffer.length > 0 && 
                        buffer[0] === 0x50 && 
                        buffer[1] === 0x4B; // PK (ZIP header)
    
    if (isValidExcel) {
        console.log('✅ Archivo Excel válido (formato ZIP/XLSX)\n');
        testResults.push({ name: 'Formato Excel', status: 'PASSED' });
    } else {
        console.log('❌ Archivo no es un Excel válido\n');
        testResults.push({ name: 'Formato Excel', status: 'FAILED' });
        allTestsPassed = false;
    }
}

// 6. Test de linting
try {
    console.log('🔍 Ejecutando linting...');
    execSync('npm run lint', { stdio: 'inherit' });
    console.log('✅ Linting pasado\n');
    testResults.push({ name: 'Linting', status: 'PASSED' });
} catch (error) {
    console.log('❌ Linting falló\n');
    testResults.push({ name: 'Linting', status: 'FAILED' });
    allTestsPassed = false;
}

// 7. Test de type checking
try {
    console.log('🔍 Ejecutando type checking...');
    execSync('npm run type-check', { stdio: 'inherit' });
    console.log('✅ Type checking pasado\n');
    testResults.push({ name: 'Type Checking', status: 'PASSED' });
} catch (error) {
    console.log('❌ Type checking falló\n');
    testResults.push({ name: 'Type Checking', status: 'FAILED' });
    allTestsPassed = false;
}

// Resumen final
console.log('📊 ========================================');
console.log('📊 RESUMEN DE TESTS');
console.log('📊 ========================================');

testResults.forEach(result => {
    const status = result.status === 'PASSED' ? '✅' : '❌';
    const duration = result.duration ? ` (${result.duration}ms)` : '';
    const fileSize = result.fileSize ? ` (${result.fileSize} KB)` : '';
    console.log(`${status} ${result.name}${duration}${fileSize}`);
});

console.log('\n📈 ========================================');
console.log('📈 ESTADÍSTICAS');
console.log('📈 ========================================');

const passed = testResults.filter(r => r.status === 'PASSED').length;
const total = testResults.length;
const percentage = ((passed / total) * 100).toFixed(1);

console.log(`✅ Tests pasados: ${passed}/${total} (${percentage}%)`);
console.log(`❌ Tests fallidos: ${total - passed}`);

if (allTestsPassed) {
    console.log('\n🎉 ¡TODOS LOS TESTS PASARON!');
    console.log('🚀 Han Excel Builder está listo para usar');
    process.exit(0);
} else {
    console.log('\n💥 ALGUNOS TESTS FALLARON');
    console.log('🔧 Revisa los errores arriba');
    process.exit(1);
} 
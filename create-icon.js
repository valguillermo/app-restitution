#!/usr/bin/env node

/**
 * Script pour créer automatiquement icon.ico à partir de icon.svg
 * Exécutez avec: node create-icon.js
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');

console.log('\n' + '='.repeat(60));
console.log('🎨 Générateur d\'Icône - BUT Restitution');
console.log('='.repeat(60) + '\n');

const iconSvgPath = path.join(__dirname, 'icon.svg');
const iconIcoPath = path.join(__dirname, 'icon.ico');
const iconPngPath = path.join(__dirname, 'icon.png');

// Vérifier si icon.svg existe
if (!fs.existsSync(iconSvgPath)) {
    console.error('❌ Erreur: icon.svg non trouvé dans le dossier courant');
    process.exit(1);
}

console.log('📝 Étape 1: Installation de sharp...\n');

try {
    // Installer sharp
    execSync('npm install --save-dev sharp', { stdio: 'inherit' });
    console.log('\n✅ Sharp installé avec succès!\n');
} catch (err) {
    console.error('❌ Erreur lors de l\'installation de sharp');
    process.exit(1);
}

console.log('🎨 Étape 2: Conversion SVG → PNG...\n');

try {
    const sharp = require('sharp');

    // Convertir SVG en PNG
    sharp(iconSvgPath)
        .resize(256, 256, {
            fit: 'contain',
            background: { r: 102, g: 126, b: 234, alpha: 1 }
        })
        .png()
        .toFile(iconPngPath)
        .then(info => {
            console.log('✅ PNG créé avec succès!');
            console.log(`   Dimensions: ${info.width}x${info.height}\n`);
            console.log('📦 Étape 3: Conversion PNG → ICO...\n');

            // Créer l'ICO à partir du PNG
            createIco(iconPngPath, iconIcoPath);
        })
        .catch(err => {
            console.error('❌ Erreur lors de la conversion SVG:', err.message);
            process.exit(1);
        });

} catch (err) {
    console.error('❌ Erreur:', err.message);
    process.exit(1);
}

function createIco(pngPath, icoPath) {
    try {
        const sharp = require('sharp');

        // Lire le PNG et le convertir en ICO
        sharp(pngPath)
            .resize(256, 256)
            .toBuffer()
            .then(buffer => {
                // Créer un ICO valide avec les dimensions correctes
                const icoBuffer = createIcoFromBuffer(buffer, 256, 256);
                fs.writeFileSync(icoPath, icoBuffer);

                console.log('✅ ICO créé avec succès!');
                console.log(`   Fichier: ${path.basename(icoPath)}\n`);

                // Nettoyer le PNG temporaire
                if (fs.existsSync(iconPngPath)) {
                    fs.unlinkSync(iconPngPath);
                    console.log('🧹 Fichier temporaire supprimé\n');
                }

                console.log('='.repeat(60));
                console.log('✨ Icône prête!');
                console.log('='.repeat(60));
                console.log('\nVous pouvez maintenant créer l\'installateur:');
                console.log('  npm run build\n');
            })
            .catch(err => {
                console.error('❌ Erreur lors de la conversion PNG:', err.message);
                process.exit(1);
            });

    } catch (err) {
        console.error('❌ Erreur:', err.message);
        process.exit(1);
    }
}

function createIcoFromBuffer(buffer, width, height) {
    // Créer un header ICO simple (256x256)
    // Format: ICONDIR structure suivi de ICONDIRENTRY et données bitmap

    const icoHeader = Buffer.alloc(6);
    icoHeader.writeUInt16LE(0, 0);      // Réservé
    icoHeader.writeUInt16LE(1, 2);      // Type: 1 = ICO
    icoHeader.writeUInt16LE(1, 4);      // Nombre d'images

    const icoEntry = Buffer.alloc(16);
    icoEntry.writeUInt8(width, 0);      // Largeur
    icoEntry.writeUInt8(height, 1);     // Hauteur
    icoEntry.writeUInt8(0, 2);          // Nombre de couleurs
    icoEntry.writeUInt8(0, 3);          // Réservé
    icoEntry.writeUInt16LE(1, 4);       // Plans de couleur
    icoEntry.writeUInt16LE(32, 6);      // Bits par pixel
    icoEntry.writeUInt32LE(buffer.length, 8);  // Taille de l'image
    icoEntry.writeUInt32LE(22, 12);     // Offset des données

    // Combiner tous les buffers
    return Buffer.concat([icoHeader, icoEntry, buffer]);
}

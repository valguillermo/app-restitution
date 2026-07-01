// Configuration Admin
const ADMIN_PASSWORD = 'BUT2026';
let adminAuthenticated = false;

// État de l'application
let appState = {
    transporter: {},
    data: {
        clients: [],
        products: {},
        submissions: []
    },
    currentClient: null,
    currentProduct: null,
    loadedFile: {
        name: '',
        period: '',
        loadDate: ''
    }
};

// Vérifier si localStorage est disponible
let useLocalStorage = true;
try {
    localStorage.setItem('test', 'test');
    localStorage.removeItem('test');
} catch(e) {
    useLocalStorage = false;
    console.warn('localStorage non disponible, utilisation de la mémoire');
}

// Charger l'état depuis le localStorage
function loadState() {
    try {
        if (useLocalStorage) {
            const saved = localStorage.getItem('appState');
            if (saved) {
                appState = JSON.parse(saved);
                loadTransporter();
            }
        }
    } catch(e) {
        console.warn('Erreur loadState:', e);
    }
}

// Sauvegarder l'état
function saveState() {
    try {
        if (useLocalStorage) {
            localStorage.setItem('appState', JSON.stringify(appState));
        }
    } catch(e) {
        console.warn('Erreur saveState:', e);
    }
}

// Changer d'onglet
function switchTab(tabName) {
    if ((tabName === 'restitution' || tabName === 'export') && (!appState.loadedFile || !appState.loadedFile.name)) {
        alert('⚠️ Aucun fichier chargé!\n\nL\'administrateur doit d\'abord charger un fichier Excel avant que vous puissiez enregistrer les restitutions.');
        return;
    }

    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });

    document.getElementById(tabName).classList.add('active');
    if (event && event.target) {
        event.target.classList.add('active');
    }
}

// Sauvegarder les informations du transporteur
function saveTransporter() {
    appState.transporter = {
        nom: document.getElementById('nom').value,
        prenom: document.getElementById('prenom').value,
        societe: document.getElementById('societe').value
    };
    saveState();
}

// Charger les informations du transporteur
function loadTransporter() {
    if (appState.transporter) {
        document.getElementById('nom').value = appState.transporter.nom || '';
        document.getElementById('prenom').value = appState.transporter.prenom || '';
        document.getElementById('societe').value = appState.transporter.societe || '';
    }
}

// Gérer le drag & drop
const uploadArea = document.getElementById('uploadArea');
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});
uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});
uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) {
        document.getElementById('fileInput').files = e.dataTransfer.files;
        loadFile({target: {files: e.dataTransfer.files}});
    }
});

// Charger le fichier Excel
function loadFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    if (appState.loadedFile.name && !confirm('Un fichier est déjà chargé. Les données de restitution seront réinitialisées. Continuer?')) {
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array', defval: '' });
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '', blankrows: false });

            if (rows.length === 0) {
                alert('❌ Le fichier est vide ou n\'a pas de données lisibles.\n\nVérifiez que:\n- Le fichier contient des données\n- Vous avez téléchargé le bon fichier');
                return;
            }

            console.log('Colonnes trouvées:', Object.keys(rows[0]));
            parseExcelData(rows);

            if (appState.data.clients.length === 0) {
                alert('❌ Aucun client trouvé dans le fichier.\n\nVérifiez que le fichier contient les colonnes suivantes:\n- OPL\n- Nom\n- Prénom\n- Désignation produit\n- SKU Evollis\n- Catégorie');
                return;
            }

            appState.loadedFile = {
                name: file.name,
                period: extractPeriod(file.name),
                loadDate: new Date().toLocaleString('fr-FR')
            };

            document.getElementById('fileStatus').classList.remove('hidden');
            document.getElementById('fileInfo').innerHTML = `
                ✅ Fichier chargé avec succès!<br>
                <strong>${file.name}</strong><br>
                Période: ${appState.loadedFile.period}<br>
                ${rows.length} enregistrements trouvés<br>
                ${appState.data.clients.length} clients
            `;

            updateFileInfo();
            saveState();

            setTimeout(() => {
                alert('✅ Fichier chargé avec succès!\n\nClients: ' + appState.data.clients.length + '\nProduits: ' + appState.data.clients.reduce((sum, c) => sum + c.products.length, 0) + '\n\nLes transporteurs peuvent maintenant utiliser l\'outil.');
            }, 500);
        } catch (error) {
            console.error('Erreur complète:', error);
            alert('❌ Erreur lors de la lecture du fichier:\n\n' + error.message + '\n\nVérifiez que le fichier est au format .xlsx correct.');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Fonction pour trouver une colonne
function findColumn(row, variants) {
    for (let variant of variants) {
        if (row.hasOwnProperty(variant) && row[variant] !== '' && row[variant] !== null && row[variant] !== undefined) {
            return row[variant];
        }
    }
    return '';
}

// Parser les données Excel
function parseExcelData(rows) {
    const clientsMap = new Map();

    rows.forEach((row, index) => {
        const opl = findColumn(row, ['OPL', 'opl', 'Opl']);
        const nomClient = findColumn(row, ['Nom', 'nom']);
        const prenomClient = findColumn(row, ['Prénom', 'prenom']);
        const designation = findColumn(row, ['Désignation produit', 'designation produit']);
        const sku = findColumn(row, ['SKU Evollis', 'sku evollis']);
        const category = findColumn(row, ['Catégorie (Sous catégorie)', 'Catégorie', 'categorie']);

        if (!opl || !designation || !sku) {
            return;
        }

        const clientKey = `${opl}-${nomClient}`;
        if (!clientsMap.has(clientKey)) {
            clientsMap.set(clientKey, {
                opl: opl,
                nom: nomClient,
                prenom: prenomClient,
                products: []
            });
        }

        clientsMap.get(clientKey).products.push({
            id: `${opl}-${sku}`,
            designation: designation.toString().trim(),
            sku: sku.toString().trim(),
            category: category.toString().trim(),
            checked: false,
            grade: '',
            motif: '',
            photos: []
        });
    });

    appState.data.clients = Array.from(clientsMap.values());
    appState.data.products = {};

    populateClientSelect();
    saveState();
}

// Vérifier si un client a des produits non cochés
function clientHasUncheckedProducts(client) {
    return client.products.some(p => !p.checked);
}

// Remplir la liste des clients
function populateClientSelect() {
    const select = document.getElementById('clientSelect');
    const currentValue = select.value;
    select.innerHTML = '<option value="">-- Choisir un client --</option>';

    appState.data.clients.forEach(client => {
        if (clientHasUncheckedProducts(client)) {
            const option = document.createElement('option');
            option.value = client.opl;
            option.textContent = `${client.opl} - ${client.nom} ${client.prenom}`;
            select.appendChild(option);
        }
    });

    if (currentValue && clientHasUncheckedProducts(appState.data.clients.find(c => c.opl === currentValue))) {
        select.value = currentValue;
    } else {
        select.value = '';
    }
}

// Charger les produits du client
function loadClientProducts() {
    const opl = document.getElementById('clientSelect').value;
    if (!opl) {
        document.getElementById('clientDetails').classList.add('hidden');
        return;
    }

    const client = appState.data.clients.find(c => c.opl === opl);
    appState.currentClient = client;

    document.getElementById('clientInfoDisplay').textContent =
        `${client.opl} - ${client.nom} ${client.prenom}`;
    document.getElementById('clientDetails').classList.remove('hidden');

    displayProducts(client.products);
    updateRestitutionSummary();
    saveState();
}

// Mettre à jour le résumé
function updateRestitutionSummary() {
    if (!appState.currentClient) return;

    const checked = appState.currentClient.products.filter(p => p.checked).length;
    const unchecked = appState.currentClient.products.filter(p => !p.checked).length;
    const total = appState.currentClient.products.length;

    document.getElementById('checkedCount').textContent = checked;
    document.getElementById('uncheckedCount').textContent = unchecked;
    document.getElementById('totalCount').textContent = total;
}

// Afficher les produits
function displayProducts(products) {
    const container = document.getElementById('productsList');
    container.innerHTML = '';

    const needsGrade = ['cuisinière', 'four', 'table de cuisson', 'cave à vin', 'combiné',
                      'congélateur', 'réfrigérateur', 'réfrigérateur US', 'lave linge',
                      'lave vaisselle', 'seche linge', 'téléviseur', 'accessoire image'];

    products.forEach((product, index) => {
        const needsGradeForProduct = needsGrade.some(cat =>
            product.category?.toLowerCase().includes(cat.toLowerCase())
        );

        const productDiv = document.createElement('div');
        productDiv.className = `product-item ${product.checked ? 'checked' : ''}`;

        productDiv.innerHTML = `
            <div class="product-header">
                <input type="checkbox" ${product.checked ? 'checked' : ''}
                       onchange="toggleProduct(${index}, this)">
                <div class="product-info">
                    <div class="product-designation">${product.designation}</div>
                    <div class="product-sku">SKU: ${product.sku}</div>
                    <div class="product-category">${product.category || 'Non catégorisé'}</div>
                </div>
            </div>
            <div class="product-details">
                ${needsGradeForProduct ? `
                    <div>
                        <strong>Grade:</strong> ${product.grade ? `<span style="color: #28a745;">${product.grade}</span>` : '<span style="color: #dc3545;">Non renseigné</span>'}
                    </div>
                    <div>
                        <strong>Motif:</strong> ${product.motif ? `<span style="color: #28a745;">${product.motif}</span>` : '<span style="color: #dc3545;">Non renseigné</span>'}
                    </div>
                ` : '<em style="color: #999;">Aucun détail supplémentaire requis</em>'}
                ${product.photos && product.photos.length > 0 ? `
                    <div><strong>Photos:</strong> ${product.photos.length} fichier(s)</div>
                ` : ''}
                ${needsGradeForProduct ? `<button class="btn btn-small" onclick="openProductModal(${index})">Détails</button>` : ''}
            </div>
        `;

        container.appendChild(productDiv);
    });
}

// Basculer le produit
function toggleProduct(index, checkbox) {
    if (!appState.currentClient) return;

    const needsGrade = ['cuisinière', 'four', 'table de cuisson', 'cave à vin', 'combiné',
                      'congélateur', 'réfrigérateur', 'réfrigérateur US', 'lave linge',
                      'lave vaisselle', 'seche linge', 'téléviseur', 'accessoire image'];

    const product = appState.currentClient.products[index];
    const needsGradeForProduct = needsGrade.some(cat =>
        product.category?.toLowerCase().includes(cat.toLowerCase())
    );

    if (checkbox.checked && needsGradeForProduct) {
        appState.currentProduct = index;
        openProductModal(index);
    } else {
        product.checked = checkbox.checked;
        if (!checkbox.checked) {
            product.grade = '';
            product.motif = '';
            product.photos = [];
        }
        displayProducts(appState.currentClient.products);
        updateRestitutionSummary();
        populateClientSelect();
        saveState();
    }
}

// Ouvrir le modal du produit
function openProductModal(index) {
    const product = appState.currentClient.products[index];
    appState.currentProduct = index;

    document.getElementById('modalTitle').textContent = product.designation;
    document.getElementById('gradeSelect').value = product.grade || '';
    document.getElementById('motifSelect').value = product.motif || '';
    document.getElementById('photosUpload').value = '';

    displayPhotosList(product.photos || []);
    toggleMotifSection();
    document.getElementById('productModal').style.display = 'block';
}

// Fermer le modal
function closeProductModal() {
    document.getElementById('productModal').style.display = 'none';
}

// Gérer l'upload de photos
function handlePhotoUpload() {
    const files = document.getElementById('photosUpload').files;
    if (!files.length || appState.currentProduct === null) return;

    const product = appState.currentClient.products[appState.currentProduct];
    if (!product.photos) product.photos = [];

    Array.from(files).forEach(file => {
        const reader = new FileReader();
        reader.onload = function(e) {
            product.photos.push({
                name: file.name,
                data: e.target.result
            });
            displayPhotosList(product.photos);
        };
        reader.readAsDataURL(file);
    });
}

// Afficher la liste des photos
function displayPhotosList(photos) {
    const container = document.getElementById('photosList');
    container.innerHTML = '';

    photos.forEach((photo, index) => {
        const div = document.createElement('div');
        div.style.marginBottom = '10px';
        div.innerHTML = `
            <div style="font-size: 12px; color: #666;">${photo.name}</div>
            <button class="btn btn-danger btn-small" onclick="removePhoto(${index})">Supprimer</button>
        `;
        container.appendChild(div);
    });
}

// Supprimer une photo
function removePhoto(index) {
    const product = appState.currentClient.products[appState.currentProduct];
    product.photos.splice(index, 1);
    displayPhotosList(product.photos);
}

// Afficher/masquer la section des motifs
function toggleMotifSection() {
    const grade = document.getElementById('gradeSelect').value;
    const motifGroup = document.getElementById('motifGroup');
    const motifSelect = document.getElementById('motifSelect');

    if (grade === 'A' || grade === 'B') {
        motifGroup.style.display = 'none';
        motifSelect.value = '';
    } else if (grade === 'C' || grade === 'D') {
        motifGroup.style.display = 'block';
    } else {
        motifGroup.style.display = 'block';
    }
}

// Mettre à jour les données du produit
function updateProductData() {
    if (appState.currentProduct === null || !appState.currentClient) return;

    const product = appState.currentClient.products[appState.currentProduct];
    product.grade = document.getElementById('gradeSelect').value;
    product.motif = document.getElementById('motifSelect').value;

    saveState();
}

// Sauvegarder les données du produit
function saveProductData() {
    const product = appState.currentClient.products[appState.currentProduct];
    const grade = document.getElementById('gradeSelect').value;
    const motif = document.getElementById('motifSelect').value;

    const needsGrade = ['cuisinière', 'four', 'table de cuisson', 'cave à vin', 'combiné',
                      'congélateur', 'réfrigérateur', 'réfrigérateur US', 'lave linge',
                      'lave vaisselle', 'seche linge', 'téléviseur', 'accessoire image'];

    const needsGradeForProduct = needsGrade.some(cat =>
        product.category?.toLowerCase().includes(cat.toLowerCase())
    );

    if (needsGradeForProduct && !grade) {
        alert('Veuillez sélectionner un grade');
        return;
    }

    if (needsGradeForProduct && (grade === 'C' || grade === 'D') && !motif) {
        alert('Veuillez sélectionner un motif pour ce grade');
        return;
    }

    product.checked = true;
    product.grade = grade;
    product.motif = motif || '';

    saveState();
    displayProducts(appState.currentClient.products);
    updateRestitutionSummary();
    populateClientSelect();
    closeProductModal();
}

// Chercher par OPL
function searchByOPL() {
    const opl = document.getElementById('oplSearch').value.trim();
    if (opl) {
        const client = appState.data.clients.find(c => c.opl === opl);
        if (client) {
            document.getElementById('clientSelect').value = opl;
            loadClientProducts();
        } else {
            alert('OPL non trouvé');
        }
    }
}

// Exporter vers Excel
function exportToExcel() {
    if (!appState.data.submissions || appState.data.submissions.length === 0) {
        alert('❌ Aucune restitution enregistrée à exporter.\n\nVeuillez d\'abord enregistrer des restitutions.');
        return;
    }

    const exportData = [];

    appState.data.submissions.forEach(submission => {
        submission.produits.forEach(product => {
            exportData.push({
                'OPL': submission.client.opl,
                'Nom client': submission.client.nom,
                'Prénom client': submission.client.prenom,
                'Désignation produit': product.designation,
                'Sku Evollis': product.sku,
                'Grade': product.grade || '',
                'Motif': product.motif || '',
                'Date enregistrement': new Date(submission.timestamp).toLocaleString('fr-FR'),
                'Transporteur': `${submission.transporteur.nom} ${submission.transporteur.prenom}`,
                'Société': submission.transporteur.societe
            });
        });
    });

    if (exportData.length === 0) {
        alert('Aucun produit enregistré pour l\'export');
        return;
    }

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Restitutions');

    const date = new Date().toISOString().slice(0, 10);
    const period = appState.loadedFile.period || 'données';
    XLSX.writeFile(workbook, `restitutions_${period}_${date}.xlsx`);

    alert('✅ Export réussi!\n\n' + exportData.length + ' produit(s) enregistré(s)');
}

// Réinitialiser les données
function clearAllData() {
    if (confirm('Êtes-vous sûr de vouloir réinitialiser toutes les données? Cette action est irréversible.')) {
        localStorage.clear();
        location.reload();
    }
}

// Générer le résumé d'export
function generateExportSummary() {
    const container = document.getElementById('exportSummary');
    let totalSubmissions = 0;
    let totalProducts = 0;

    if (appState.data.submissions && appState.data.submissions.length > 0) {
        totalSubmissions = appState.data.submissions.length;
        totalProducts = appState.data.submissions.reduce((sum, s) => sum + s.produits.length, 0);
    }

    container.innerHTML = `
        <div class="alert alert-info" style="margin-bottom: 20px;">
            📊 <strong>Données prêtes pour l'export</strong>
        </div>
        <table class="summary-table">
            <tr>
                <td>Restitutions enregistrées</td>
                <td><strong style="color: #667eea;">${totalSubmissions}</strong></td>
            </tr>
            <tr>
                <td>Produits restitués</td>
                <td><strong style="color: #28a745;">${totalProducts}</strong></td>
            </tr>
            <tr>
                <td>Fichier actuellement chargé</td>
                <td><strong>${appState.loadedFile.name ? appState.loadedFile.name : 'Aucun'}</strong></td>
            </tr>
        </table>
    `;
}

// Valider la restitution
function validateRestitution() {
    if (!appState.currentClient) {
        alert('Sélectionnez d\'abord un client');
        return;
    }

    const checkedProducts = appState.currentClient.products.filter(p => p.checked);
    if (checkedProducts.length === 0) {
        alert('Sélectionnez au moins un produit à restituer');
        return;
    }

    const needsGrade = ['cuisinière', 'four', 'table de cuisson', 'cave à vin', 'combiné',
                      'congélateur', 'réfrigérateur', 'réfrigérateur US', 'lave linge',
                      'lave vaisselle', 'seche linge', 'téléviseur', 'accessoire image'];

    for (let product of checkedProducts) {
        const needsGradeForProduct = needsGrade.some(cat =>
            product.category?.toLowerCase().includes(cat.toLowerCase())
        );

        if (needsGradeForProduct && !product.grade) {
            alert(`❌ Le produit "${product.designation}" nécessite un grade.\n\nVeuillez compléter les informations.`);
            return;
        }

        if (needsGradeForProduct && (product.grade === 'C' || product.grade === 'D') && !product.motif) {
            alert(`❌ Le produit "${product.designation}" (Grade ${product.grade}) nécessite un motif.\n\nVeuillez compléter les informations.`);
            return;
        }
    }

    showValidationModal(checkedProducts);
}

// Afficher le modal de validation
function showValidationModal(products) {
    const transporteur = `${document.getElementById('nom').value} ${document.getElementById('prenom').value} (${document.getElementById('societe').value})`;
    const dateTime = new Date().toLocaleString('fr-FR');

    document.getElementById('validationTransporter').textContent = transporteur;
    document.getElementById('validationDateTime').textContent = dateTime;

    const table = document.getElementById('validationTable');
    table.innerHTML = `
        <tr style="background: #667eea; color: white;">
            <th>Produit</th>
            <th>SKU</th>
            <th>Catégorie</th>
            <th>Grade</th>
            <th>Motif</th>
        </tr>
    `;

    products.forEach(product => {
        const row = table.insertRow();
        row.innerHTML = `
            <td><strong>${product.designation}</strong></td>
            <td>${product.sku}</td>
            <td>${product.category}</td>
            <td>${product.grade || '-'}</td>
            <td>${product.motif || '-'}</td>
        `;
    });

    document.getElementById('validationModal').style.display = 'block';
}

// Fermer le modal de validation
function closeValidationModal() {
    document.getElementById('validationModal').style.display = 'none';
}

// Confirmer et enregistrer la restitution
function confirmRestitution() {
    if (!appState.currentClient) return;

    const restitution = {
        id: Date.now(),
        timestamp: new Date().toISOString(),
        transporteur: {
            nom: document.getElementById('nom').value,
            prenom: document.getElementById('prenom').value,
            societe: document.getElementById('societe').value
        },
        client: {
            opl: appState.currentClient.opl,
            nom: appState.currentClient.nom,
            prenom: appState.currentClient.prenom
        },
        produits: appState.currentClient.products.filter(p => p.checked)
    };

    if (!appState.data.submissions) {
        appState.data.submissions = [];
    }
    appState.data.submissions.push(restitution);

    saveState();
    closeValidationModal();

    alert(`✅ Restitution validée avec succès!\n\n${restitution.produits.length} produit(s) enregistré(s)\n\nClient: ${appState.currentClient.nom} ${appState.currentClient.prenom}\nOPL: ${appState.currentClient.opl}`);

    document.getElementById('clientSelect').value = '';
    document.getElementById('oplSearch').value = '';
    document.getElementById('clientDetails').classList.add('hidden');

    generateExportSummary();
}

// Afficher le login admin
function showAdminLogin() {
    if (adminAuthenticated) {
        switchTabAdmin('admin');
    } else {
        document.getElementById('adminLoginModal').style.display = 'block';
        document.getElementById('adminPassword').value = '';
        document.getElementById('adminPassword').focus();
    }
}

// Fermer le login admin
function closeAdminLogin() {
    document.getElementById('adminLoginModal').style.display = 'none';
}

// Vérifier le mot de passe admin
function verifyAdminPassword() {
    const password = document.getElementById('adminPassword').value;
    if (password === ADMIN_PASSWORD) {
        adminAuthenticated = true;
        closeAdminLogin();
        switchTabAdmin('admin');
        updateFileInfo();
    } else {
        alert('Mot de passe incorrect');
        document.getElementById('adminPassword').value = '';
    }
}

// Changer d'onglet (version admin)
function switchTabAdmin(tabName) {
    document.querySelectorAll('.tab-content').forEach(tab => {
        tab.classList.remove('active');
    });
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });

    document.getElementById(tabName).classList.add('active');

    const adminBtns = Array.from(document.querySelectorAll('.tab-btn'));
    const adminBtn = adminBtns.find(btn => btn.textContent.includes('⚙️'));
    if (adminBtn) {
        adminBtn.classList.add('active');
    }
}

// Mettre à jour les infos du fichier chargé
function updateFileInfo() {
    const infoDiv = document.getElementById('currentFileInfo');
    if (appState.loadedFile && appState.loadedFile.name) {
        infoDiv.innerHTML = `
            <div>
                <strong>Fichier:</strong> ${appState.loadedFile.name}<br>
                <strong>Période:</strong> ${appState.loadedFile.period || 'Période détectée'}<br>
                <strong>Date de chargement:</strong> ${appState.loadedFile.loadDate}<br>
                <strong>Nombre de clients:</strong> ${appState.data.clients.length}<br>
                <strong>Nombre de produits:</strong> ${appState.data.clients.reduce((sum, c) => sum + c.products.length, 0)}
            </div>
        `;
    } else {
        infoDiv.innerHTML = '<em style="color: #999;">Aucun fichier chargé - Les transporteurs ne peuvent pas utiliser l\'outil</em>';
    }
}

// Exporter avant de charger un nouveau fichier
function exportBeforeNewFile() {
    const checkedProducts = [];
    appState.data.clients.forEach(client => {
        client.products.forEach(product => {
            if (product.checked) {
                checkedProducts.push({
                    'OPL': client.opl,
                    'Nom client': client.nom,
                    'Prénom client': client.prenom,
                    'Désignation produit': product.designation,
                    'Sku Evollis': product.sku,
                    'Grade': product.grade || '',
                    'Motif': product.motif || ''
                });
            }
        });
    });

    if (checkedProducts.length === 0) {
        alert('Aucune donnée à exporter');
        return;
    }

    const worksheet = XLSX.utils.json_to_sheet(checkedProducts);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Restitutions');

    const date = new Date().toISOString().slice(0, 10);
    const period = appState.loadedFile.period || 'données';
    XLSX.writeFile(workbook, `restitutions_${period}_${date}.xlsx`);

    alert('Export terminé. Vous pouvez maintenant charger un nouveau fichier.');
}

// Extraire la période du nom du fichier
function extractPeriod(filename) {
    const match = filename.match(/reporting_but_fin_de_contrat_([T|Q][\d]_[\d]{4})/i);
    return match ? match[1] : 'Inconnue';
}

// Mettre à jour l'affichage du statut du fichier
function updateFileStatusDisplay() {
    const fileStatusAlert = document.getElementById('fileStatusAlert');
    const noFileAlert = document.getElementById('noFileAlert');
    const fileStatusAlertText = document.getElementById('fileStatusAlertText');

    if (appState.loadedFile && appState.loadedFile.name) {
        fileStatusAlert.style.display = 'block';
        noFileAlert.style.display = 'none';
        fileStatusAlertText.innerHTML = `
            Fichier: <strong>${appState.loadedFile.name}</strong><br>
            Période: ${appState.loadedFile.period}<br>
            Chargé le: ${appState.loadedFile.loadDate}
        `;
    } else {
        fileStatusAlert.style.display = 'none';
        noFileAlert.style.display = 'block';
    }
}

// Fermer le modal en cliquant en dehors
window.onclick = function(event) {
    const modal = document.getElementById('productModal');
    if (event.target === modal) {
        modal.style.display = 'none';
    }
};

// Initialiser l'application
loadState();
updateFileInfo();
updateFileStatusDisplay();

setInterval(generateExportSummary, 1000);
setInterval(updateFileStatusDisplay, 1000);

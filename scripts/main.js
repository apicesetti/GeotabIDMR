/**
 * Movimientos de equipos — add-in MyGeotab.
 * Registra instalaciones, desinstalaciones y reemplazos declarados por el cliente.
 * Solo escribe en AddInData (addInId fijo): nunca modifica Device.
 */
geotab.addin.movimientos = function () {
    'use strict';

    var ADDIN_ID = 'aOS_yCVsbRPiyLHnQhiSteg';
    var RECORD_TYPE = 'movimiento';
    var HARNESSES = ['HRN-CW03', 'HRN-GS09', 'HRN-GS16K2-HD', 'HRN-GS16K2-LD', 'HRN-GS16K22'];
    var CAMERAS = ['GO FOCUS PLUS', 'GO FOCUS PRO', 'SURFSIGHT AI12'];
    var TYPE_LABEL = { instalacion: 'Instalación', desinstalacion: 'Desinstalación', reemplazo: 'Reemplazo' };
    var REMOVED_LABEL = { go: 'Equipo GO', go_camara: 'Equipo GO y cámara', camara: 'Solo cámara' };
    var MAX_KM = 9999999;
    var ADMIN_SECURITY_ID = 'GroupEverythingSecurityId';

    var api, root;
    var userName = '';
    var isAdmin = false;        // perfil Administrador: único que puede anular
    var devices = [];
    var records = [];
    var selected = null;        // vehículo elegido en Desinstalación / Reemplazo
    var selectToken = 0;
    var loadToken = 0;
    var warnedInstall = null;   // firma del formulario ya advertido por duplicado
    var voidTarget = null;
    var saving = false;
    var bound = false;

    // ------------------------------------------------------------------
    // Utilidades
    // ------------------------------------------------------------------
    function $(sel) { return root.querySelector(sel); }
    function $$(sel, el) { return Array.prototype.slice.call((el || root).querySelectorAll(sel)); }

    function esc(s) {
        return String(s === undefined || s === null ? '' : s)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }
    function up(v) { return String(v || '').trim().toUpperCase(); }
    function norm(v) { return up(v).replace(/[^A-Z0-9]/g, ''); }
    function pad(n) { return (n < 10 ? '0' : '') + n; }
    function today() {
        var d = new Date();
        return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
    }
    function fmtDay(s) {
        var p = String(s || '').split('-');
        return p.length === 3 ? p[2] + '/' + p[1] + '/' + p[0] : '';
    }
    function fmtKm(v) {
        return v === null || v === undefined || v === '' ? '' : Number(v).toLocaleString('es-AR') + ' km';
    }
    function fmtDateTime(iso) {
        var d = new Date(iso);
        return !iso || isNaN(d) ? '' : d.toLocaleString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
    }
    function options(list, placeholder, value) {
        var html = '<option value="">' + esc(placeholder) + '</option>';
        list.forEach(function (o) {
            html += '<option value="' + esc(o) + '"' + (o === value ? ' selected' : '') + '>' + esc(o) + '</option>';
        });
        return html;
    }
    function serialOf(d) {
        return d && !/^0*$/.test(norm(d.serialNumber)) ? d.serialNumber : '';
    }
    function isArchived(d) {
        return !!d.activeTo && new Date(d.activeTo) < new Date();
    }
    function call(method, params) {
        return new Promise(function (resolve, reject) { api.call(method, params, resolve, reject); });
    }
    function errText(e) {
        return (e && (e.message || (e.responseJSON && e.responseJSON.error && e.responseJSON.error.message))) || String(e || 'error desconocido');
    }

    function showMsg(text, kind) {
        var el = $('#mv-msg');
        el.textContent = text;
        el.className = 'mv-msg mv-msg-' + (kind || 'ok');
        el.hidden = false;
        if (root.scrollIntoView) { root.scrollIntoView({ block: 'start' }); }
    }
    function clearMsg() { $('#mv-msg').hidden = true; }

    function clearErrors(form) {
        $$('.mv-error', form).forEach(function (n) { n.parentNode.removeChild(n); });
        $$('.is-invalid', form).forEach(function (n) { n.classList.remove('is-invalid'); });
    }
    function fieldError(form, name, text) {
        var input = form.elements[name];
        input.classList.add('is-invalid');
        var span = document.createElement('span');
        span.className = 'mv-error';
        span.textContent = text;
        input.parentNode.appendChild(span);
    }

    /** Valida odómetro y fecha comunes; devuelve true si están bien. */
    function checkCommon(form) {
        var ok = true, km = form.elements.odometer.value;
        if (km === '' || !/^\d+$/.test(km) || Number(km) > MAX_KM) {
            fieldError(form, 'odometer', 'Ingresá los km como número entero, sin puntos.');
            ok = false;
        }
        var date = form.elements.date.value;
        if (!date) { fieldError(form, 'date', 'Indicá la fecha.'); ok = false; }
        else if (date > today()) { fieldError(form, 'date', 'La fecha no puede ser futura.'); ok = false; }
        return ok;
    }
    function goSerialValid(v) { return /^[A-Z0-9]{12}$/.test(v) && !/^0+$/.test(v); }

    // ------------------------------------------------------------------
    // Datos
    // ------------------------------------------------------------------
    function loadData() {
        var token = ++loadToken;
        return Promise.all([
            call('Get', {
                typeName: 'Device',
                propertySelector: { fields: ['id', 'name', 'serialNumber', 'licensePlate', 'vehicleIdentificationNumber', 'activeTo'] }
            }),
            call('Get', { typeName: 'AddInData', search: { addInId: ADDIN_ID } }),
            loadAdmin()
        ]).then(function (res) {
            if (token !== loadToken) { return; }
            devices = res[0] || [];
            records = (res[1] || []).filter(function (r) { return r.details && r.details.type === RECORD_TYPE; });
            renderHistory();
            if (selected) { renderVehicle(); }
        }).catch(function (e) {
            showMsg('No se pudieron cargar los datos: ' + errText(e), 'err');
        });
    }

    /** Ante cualquier duda (error de API, usuario sin datos) no da permiso. */
    function loadAdmin() {
        if (!userName) { isAdmin = false; return Promise.resolve(); }
        return call('Get', { typeName: 'User', search: { name: userName } }).then(function (list) {
            var user = list && list[0];
            isAdmin = !!user && (user.securityGroups || []).some(function (g) { return g.id === ADMIN_SECURITY_ID; });
        }).catch(function () {
            isAdmin = false;
        });
    }

    function saveNew(details) {
        var entity = { addInId: ADDIN_ID, groups: [{ id: 'GroupCompanyId' }], details: details };
        return call('Add', { typeName: 'AddInData', entity: entity }).then(function (id) {
            entity.id = id;
            records.push(entity);
            renderHistory();
        });
    }

    function baseRecord(type, date, odometer, notes) {
        return {
            type: RECORD_TYPE, schemaVersion: 1, movementType: type,
            date: date, odometerKm: Number(odometer), notes: notes || '',
            createdAt: new Date().toISOString(), createdBy: userName,
            voided: false
        };
    }

    function byChrono(a, b) {
        var x = a.details, y = b.details;
        if (x.date !== y.date) { return x.date < y.date ? -1 : 1; }
        return String(x.createdAt) < String(y.createdAt) ? -1 : 1;
    }

    /** Movimientos vigentes de un vehículo, en orden cronológico. */
    function recordsOf(device) {
        var plates = [norm(device.licensePlate), norm(device.name)].filter(Boolean);
        var serial = norm(serialOf(device));
        return records.filter(function (r) {
            var d = r.details;
            if (d.voided) { return false; }
            if (d.deviceId) { return d.deviceId === device.id; }
            return plates.indexOf(norm(d.licensePlate)) >= 0 || (serial && norm(d.goSerial) === serial);
        }).sort(byChrono);
    }

    /** Estado declarado del vehículo según sus movimientos. */
    function declaredState(recs) {
        var s = { harness: '', camera: null, brand: '', model: '', chassis: '', last: null };
        recs.forEach(function (r) {
            var d = r.details;
            s.last = d;
            if (d.vehicleBrand) { s.brand = d.vehicleBrand; }
            if (d.vehicleModel) { s.model = d.vehicleModel; }
            if (d.chassis) { s.chassis = d.chassis; }
            if (d.movementType === 'instalacion') {
                s.harness = d.harness || '';
                s.camera = d.camera || null;
            } else if (d.movementType === 'desinstalacion') {
                if (d.removed !== 'go') { s.camera = null; }
                if (d.removed !== 'camara') { s.harness = ''; }
            } else if (d.replaced) {
                if (d.replaced.harness) { s.harness = d.replaced.harness.to; }
                if (d.replaced.camera) { s.camera = d.replaced.camera.to; }
            }
        });
        return s;
    }

    // ------------------------------------------------------------------
    // Instalación
    // ------------------------------------------------------------------
    function resetInstall() {
        var f = $('#mv-install');
        f.reset();
        clearErrors(f);
        f.elements.date.value = today();
        syncCameraSerial(f.elements.cameraModel, f.elements.cameraSerial, 'Elegí primero marca y modelo');
        warnedInstall = null;
    }

    function syncCameraSerial(select, input, emptyHint) {
        input.disabled = !select.value;
        input.placeholder = select.value ? '' : emptyHint;
        if (!select.value) { input.value = ''; }
    }

    function submitInstall(e) {
        e.preventDefault();
        if (saving) { return; }
        var f = e.target, el = f.elements, ok = true;
        clearErrors(f);
        clearMsg();
        var v = {
            plate: up(el.plate.value).replace(/[\s-]/g, ''), chassis: up(el.chassis.value),
            brand: el.brand.value.trim(), model: el.model.value.trim(),
            goSerial: norm(el.goSerial.value), harness: el.harness.value,
            cameraModel: el.cameraModel.value, cameraSerial: up(el.cameraSerial.value)
        };
        if (!v.plate) { fieldError(f, 'plate', 'Ingresá la patente.'); ok = false; }
        if (!v.chassis) { fieldError(f, 'chassis', 'Ingresá el número de chasis.'); ok = false; }
        if (!v.brand) { fieldError(f, 'brand', 'Ingresá la marca.'); ok = false; }
        if (!v.model) { fieldError(f, 'model', 'Ingresá el modelo.'); ok = false; }
        if (!goSerialValid(v.goSerial)) { fieldError(f, 'goSerial', 'El número de serie de GO tiene 12 letras y números.'); ok = false; }
        if (!v.harness) { fieldError(f, 'harness', 'Elegí el tipo de arnés.'); ok = false; }
        if (v.cameraModel && !v.cameraSerial) { fieldError(f, 'cameraSerial', 'Ingresá el número de serie de la cámara o quitá la cámara.'); ok = false; }
        if (!checkCommon(f)) { ok = false; }
        if (!ok) { return; }

        var bySerial = devices.filter(function (d) { return !isArchived(d) && norm(d.serialNumber) === v.goSerial; })[0];
        var byPlate = devices.filter(function (d) { return !isArchived(d) && (norm(d.licensePlate) === norm(v.plate) || norm(d.name) === norm(v.plate)); })[0];
        var signature = JSON.stringify(v);
        var warn = [];
        if (bySerial) { warn.push('el GO ' + v.goSerial + ' ya figura en Geotab en ' + (bySerial.licensePlate || bySerial.name)); }
        if (byPlate && serialOf(byPlate) && byPlate !== bySerial) { warn.push('la patente ' + v.plate + ' ya tiene en Geotab el GO ' + serialOf(byPlate)); }
        if (warn.length && warnedInstall !== signature) {
            warnedInstall = signature;
            showMsg('Revisá los datos: ' + warn.join(' y ') + '. Si son correctos, tocá de nuevo "Registrar instalación".', 'warn');
            return;
        }

        var device = bySerial || byPlate || null;
        var rec = baseRecord('instalacion', el.date.value, el.odometer.value);
        rec.deviceId = device ? device.id : null;
        rec.deviceName = device ? device.name : '';
        rec.licensePlate = v.plate;
        rec.chassis = v.chassis;
        rec.vehicleBrand = v.brand;
        rec.vehicleModel = v.model;
        rec.goSerial = v.goSerial;
        rec.harness = v.harness;
        rec.camera = v.cameraModel ? { model: v.cameraModel, serial: v.cameraSerial } : null;

        persist(rec, f, function () {
            resetInstall();
            showMsg('Instalación registrada: ' + v.plate + ' · GO ' + v.goSerial + '.', 'ok');
        });
    }

    function persist(rec, form, done) {
        var btn = form.querySelector('button[type=submit]');
        saving = true;
        btn.disabled = true;
        saveNew(rec).then(done).catch(function (e) {
            showMsg('No se pudo registrar el movimiento: ' + errText(e), 'err');
        }).then(function () {
            saving = false;
            btn.disabled = false;
        });
    }

    // ------------------------------------------------------------------
    // Búsqueda
    // ------------------------------------------------------------------
    function submitSearch(e) {
        e.preventDefault();
        clearMsg();
        selected = null;
        selectToken++;
        $('#mv-vehicle').innerHTML = '';
        var q = norm($('#mv-search-text').value);
        var box = $('#mv-results');
        box.innerHTML = '';
        if (!q) { showMsg('Ingresá una patente o un número de serie de GO.', 'warn'); return; }

        var exact = devices.filter(function (d) {
            return norm(d.licensePlate) === q || norm(serialOf(d)) === q || norm(d.name) === q;
        });
        var found = exact.length || q.length < 3 ? exact : devices.filter(function (d) {
            return norm(d.licensePlate).indexOf(q) >= 0 || norm(serialOf(d)).indexOf(q) >= 0 || norm(d.name).indexOf(q) >= 0;
        });
        if (!found.length) {
            box.innerHTML = '<p class="mv-empty">No se encontró ningún vehículo con "' + esc($('#mv-search-text').value.trim()) +
                '" en Geotab.' + (q.length < 3 ? ' Probá con al menos 3 caracteres.' : '') + '</p>';
            return;
        }
        if (found.length === 1) { selectDevice(found[0]); return; }
        box.innerHTML = '<p class="mv-muted">' + found.length + ' coincidencias. Elegí el vehículo:</p>' + found.slice(0, 25).map(function (d) {
            return '<button type="button" class="mv-result" data-device="' + esc(d.id) + '"><span><strong>' +
                esc(d.licensePlate || d.name) + '</strong> · ' + esc(d.name) + '</span><span class="mv-muted">' +
                esc(serialOf(d) || 'Sin GO') + (isArchived(d) ? ' · Archivado' : '') + '</span></button>';
        }).join('');
    }

    function selectDevice(device) {
        var token = ++selectToken;
        selected = { device: device, odometerKm: null, odometerLoaded: false, lastSeen: '' };
        $('#mv-results').innerHTML = '';
        renderVehicle();
        var now = new Date().toISOString();
        Promise.all([
            call('Get', {
                typeName: 'StatusData',
                search: {
                    deviceSearch: { id: device.id },
                    diagnosticSearch: { id: 'DiagnosticOdometerAdjustmentId' },
                    fromDate: now, toDate: now
                },
                resultsLimit: 1
            }).catch(function () { return []; }),
            call('Get', { typeName: 'DeviceStatusInfo', search: { deviceSearch: { id: device.id } }, resultsLimit: 1 }).catch(function () { return []; })
        ]).then(function (res) {
            if (token !== selectToken || !selected) { return; }
            var odo = res[0] && res[0][0];
            selected.odometerKm = odo && typeof odo.data === 'number' ? Math.round(odo.data / 1000) : null;
            selected.lastSeen = res[1] && res[1][0] ? res[1][0].dateTime : '';
            selected.odometerLoaded = true;
            updateOdometer();
        });
    }

    function updateOdometer() {
        var km = $('#mv-veh-odo'), last = $('#mv-veh-seen'), input = $('#mv-change-form input[name=odometer]');
        if (!km) { return; }
        km.textContent = selected.odometerKm === null ? 'Sin dato' : fmtKm(selected.odometerKm);
        last.textContent = selected.lastSeen ? fmtDateTime(selected.lastSeen) : 'Sin dato';
        if (input && !input.value && selected.odometerKm !== null) { input.value = selected.odometerKm; }
    }

    function dataItem(label, value, id) {
        return '<div><dt>' + esc(label) + '</dt><dd' + (id ? ' id="' + id + '"' : '') + '>' + (value ? esc(value) : '<span class="mv-muted">Sin dato</span>') + '</dd></div>';
    }

    function renderVehicle() {
        var box = $('#mv-vehicle');
        if (!selected) { box.innerHTML = ''; return; }
        var d = selected.device, st = declaredState(recordsOf(d));
        selected.state = st;
        var brandModel = [st.brand, st.model].filter(Boolean).join(' ');
        var cam = st.camera ? st.camera.model + ' · ' + st.camera.serial : '';
        var last = st.last ? TYPE_LABEL[st.last.movementType] + ' · ' + fmtDay(st.last.date) : '';

        var alert = '';
        if (isArchived(d)) {
            alert = '<div class="mv-msg mv-msg-warn">Este vehículo está archivado en Geotab.</div>';
        } else if (st.last && st.last.movementType === 'desinstalacion' && st.last.removed !== 'camara') {
            alert = '<div class="mv-msg mv-msg-warn">Este vehículo ya tiene una desinstalación registrada el ' + esc(fmtDay(st.last.date)) + '.</div>';
        }

        box.innerHTML = alert +
            '<div class="mv-card"><h2>' + esc(d.licensePlate || d.name) + '</h2><dl class="mv-data">' +
            dataItem('Nombre en Geotab', d.name) +
            dataItem('Patente', d.licensePlate) +
            dataItem('Serie de GO', serialOf(d)) +
            dataItem('Número de chasis', d.vehicleIdentificationNumber || st.chassis) +
            dataItem('Marca y modelo', brandModel) +
            dataItem('Arnés', st.harness) +
            dataItem('Cámara', cam) +
            dataItem('Odómetro en Geotab', selected.odometerLoaded ? '' : 'Consultando…', 'mv-veh-odo') +
            dataItem('Última comunicación', selected.odometerLoaded ? '' : 'Consultando…', 'mv-veh-seen') +
            dataItem('Último movimiento', last) +
            '</dl></div>' +
            changeFormHtml(st);
        var form = $('#mv-change-form');
        form.elements.date.value = today();
        syncChangeForm();
        if (selected.odometerLoaded) { updateOdometer(); }
    }

    function changeFormHtml(st) {
        var cam = st.camera || {};
        return '<form id="mv-change-form" class="mv-card" novalidate autocomplete="off">' +
            '<h2>Registrar movimiento</h2>' +
            '<div class="mv-seg">' +
            '<label class="mv-check"><input type="radio" name="kind" value="desinstalacion" checked> Desinstalación</label>' +
            '<label class="mv-check"><input type="radio" name="kind" value="reemplazo"> Reemplazo</label>' +
            '</div>' +
            '<div class="mv-grid">' +
            '<label>Fecha *<input type="date" name="date" required></label>' +
            '<label>Odómetro al momento del movimiento (km) *<input type="number" name="odometer" min="0" step="1" inputmode="numeric" required></label>' +
            '</div>' +

            '<div data-kind="desinstalacion" class="mv-sub-block"><div class="mv-grid">' +
            '<label>Qué se retiró *<select name="removed">' +
            Object.keys(REMOVED_LABEL).map(function (k) { return '<option value="' + k + '">' + esc(REMOVED_LABEL[k]) + '</option>'; }).join('') +
            '</select></label></div>' +
            '<div class="mv-grid" data-removed-camera>' +
            '<label>Cámara retirada<select name="oldCameraModelU">' + options(CAMERAS, 'Sin dato', cam.model) + '</select></label>' +
            '<label>Serie de cámara retirada<input name="oldCameraSerialU" maxlength="30" value="' + esc(cam.serial) + '"></label>' +
            '</div></div>' +

            '<div data-kind="reemplazo" class="mv-sub-block">' +
            '<p class="mv-muted" style="margin:0 0 8px">Marcá qué se reemplazó:</p>' +
            '<label class="mv-check"><input type="checkbox" name="repGo"> Equipo GO</label>' +
            '<div class="mv-grid" data-rep="repGo">' +
            '<label>Serie de GO nuevo *<input name="newGoSerial" maxlength="12" placeholder="12 caracteres"></label></div>' +
            '<div class="mv-sub-block"><label class="mv-check"><input type="checkbox" name="repCamera"> Cámara</label>' +
            '<div class="mv-grid" data-rep="repCamera">' +
            '<label>Cámara anterior<select name="oldCameraModel">' + options(CAMERAS, 'Sin dato', cam.model) + '</select></label>' +
            '<label>Serie de cámara anterior<input name="oldCameraSerial" maxlength="30" value="' + esc(cam.serial) + '"></label>' +
            '<label>Cámara nueva *<select name="newCameraModel">' + options(CAMERAS, 'Elegí marca y modelo') + '</select></label>' +
            '<label>Serie de cámara nueva *<input name="newCameraSerial" maxlength="30" placeholder="Elegí primero marca y modelo" disabled></label>' +
            '</div></div>' +
            '<div class="mv-sub-block"><label class="mv-check"><input type="checkbox" name="repHarness"> Arnés</label>' +
            '<div class="mv-grid" data-rep="repHarness">' +
            '<label>Arnés anterior<select name="oldHarness">' + options(HARNESSES, 'Sin dato', st.harness) + '</select></label>' +
            '<label>Arnés nuevo *<select name="newHarness">' + options(HARNESSES, 'Elegí el arnés') + '</select></label>' +
            '</div></div></div>' +

            '<div class="mv-sub-block"><label>Motivo u observaciones<textarea name="notes" rows="2" maxlength="300"></textarea></label></div>' +
            '<div class="mv-actions" style="margin-top:12px"><button type="submit" class="mv-btn">Registrar movimiento</button></div>' +
            '</form>';
    }

    function syncChangeForm() {
        var f = $('#mv-change-form');
        if (!f) { return; }
        var el = f.elements, kind = el.kind.value;
        $$('[data-kind]', f).forEach(function (n) { n.hidden = n.getAttribute('data-kind') !== kind; });
        $$('[data-rep]', f).forEach(function (n) { n.hidden = !el[n.getAttribute('data-rep')].checked; });
        f.querySelector('[data-removed-camera]').hidden = el.removed.value === 'go';
        syncCameraSerial(el.newCameraModel, el.newCameraSerial, 'Elegí primero marca y modelo');
    }

    function submitChange(f) {
        if (saving || !selected) { return; }
        var el = f.elements, kind = el.kind.value, d = selected.device, st = selected.state, ok = true;
        clearErrors(f);
        clearMsg();
        if (!checkCommon(f)) { ok = false; }

        var rec = baseRecord(kind, el.date.value, el.odometer.value, el.notes.value.trim());
        rec.deviceId = d.id;
        rec.deviceName = d.name;
        rec.licensePlate = d.licensePlate || '';
        rec.chassis = d.vehicleIdentificationNumber || st.chassis || '';
        rec.vehicleBrand = st.brand;
        rec.vehicleModel = st.model;
        rec.goSerial = serialOf(d);

        if (kind === 'desinstalacion') {
            rec.removed = el.removed.value;
            rec.harness = st.harness;
            rec.camera = rec.removed === 'go' ? null : {
                model: el.oldCameraModelU.value, serial: up(el.oldCameraSerialU.value)
            };
        } else {
            rec.replaced = {};
            if (!el.repGo.checked && !el.repCamera.checked && !el.repHarness.checked) {
                showMsg('Marcá al menos un componente reemplazado.', 'warn');
                return;
            }
            if (el.repGo.checked) {
                var newGo = norm(el.newGoSerial.value);
                if (!goSerialValid(newGo)) { fieldError(f, 'newGoSerial', 'El número de serie de GO tiene 12 letras y números.'); ok = false; }
                else if (newGo === norm(serialOf(d))) { fieldError(f, 'newGoSerial', 'Es igual al GO actual.'); ok = false; }
                rec.replaced.go = { from: serialOf(d), to: newGo };
            }
            if (el.repCamera.checked) {
                if (!el.newCameraModel.value) { fieldError(f, 'newCameraModel', 'Elegí marca y modelo.'); ok = false; }
                if (!up(el.newCameraSerial.value)) { fieldError(f, 'newCameraSerial', 'Ingresá el número de serie.'); ok = false; }
                rec.replaced.camera = {
                    from: { model: el.oldCameraModel.value, serial: up(el.oldCameraSerial.value) },
                    to: { model: el.newCameraModel.value, serial: up(el.newCameraSerial.value) }
                };
            }
            if (el.repHarness.checked) {
                if (!el.newHarness.value) { fieldError(f, 'newHarness', 'Elegí el arnés nuevo.'); ok = false; }
                else if (el.newHarness.value === el.oldHarness.value) { fieldError(f, 'newHarness', 'Es igual al arnés anterior.'); ok = false; }
                rec.replaced.harness = { from: el.oldHarness.value, to: el.newHarness.value };
            }
        }
        if (!ok) { return; }

        var name = d.licensePlate || d.name;
        persist(rec, f, function () {
            renderVehicle();
            showMsg(TYPE_LABEL[kind] + ' registrada: ' + name + '.', 'ok');
        });
    }

    // ------------------------------------------------------------------
    // Historial
    // ------------------------------------------------------------------
    function camText(c) {
        return c && (c.model || c.serial) ? [c.model, c.serial].filter(Boolean).join(' ') : 'sin dato';
    }

    function detailOf(d) {
        var parts = [];
        if (d.movementType === 'instalacion') {
            parts.push('Chasis ' + d.chassis);
            parts.push([d.vehicleBrand, d.vehicleModel].filter(Boolean).join(' '));
            parts.push('Arnés ' + d.harness);
            if (d.camera) { parts.push('Cámara ' + camText(d.camera)); }
        } else if (d.movementType === 'desinstalacion') {
            parts.push('Retirado: ' + (REMOVED_LABEL[d.removed] || d.removed));
            if (d.camera) { parts.push('Cámara ' + camText(d.camera)); }
        } else if (d.replaced) {
            var r = d.replaced;
            if (r.go) { parts.push('GO ' + (r.go.from || 'sin dato') + ' → ' + r.go.to); }
            if (r.camera) { parts.push('Cámara ' + camText(r.camera.from) + ' → ' + camText(r.camera.to)); }
            if (r.harness) { parts.push('Arnés ' + (r.harness.from || 'sin dato') + ' → ' + r.harness.to); }
        }
        if (d.notes) { parts.push(d.notes); }
        return parts.filter(Boolean).join(' · ');
    }

    function filteredHistory() {
        var q = norm($('#mv-h-text').value), type = $('#mv-h-type').value;
        var from = $('#mv-h-from').value, to = $('#mv-h-to').value, withVoid = $('#mv-h-void').checked;
        return records.filter(function (r) {
            var d = r.details;
            if (d.voided && !withVoid) { return false; }
            if (type && d.movementType !== type) { return false; }
            if (from && d.date < from) { return false; }
            if (to && d.date > to) { return false; }
            if (q) {
                var hay = norm([d.licensePlate, d.deviceName, d.goSerial, d.chassis, detailOf(d)].join(' '));
                if (hay.indexOf(q) < 0) { return false; }
            }
            return true;
        }).sort(byChrono).reverse();
    }

    function renderHistory() {
        var rows = filteredHistory();
        $('#mv-history').innerHTML = rows.map(function (r) {
            var d = r.details;
            return '<tr class="' + (d.voided ? 'is-void' : '') + '">' +
                '<td>' + esc(fmtDay(d.date)) + '</td>' +
                '<td><span class="mv-badge mv-badge-' + esc(d.movementType) + '">' + esc(TYPE_LABEL[d.movementType] || d.movementType) + '</span></td>' +
                '<td class="mv-strike">' + esc(d.licensePlate || d.deviceName) + '</td>' +
                '<td class="mv-strike">' + esc(d.goSerial) + '</td>' +
                '<td><span class="mv-strike">' + esc(detailOf(d)) + '</span>' +
                (d.voided ? '<br>Anulado por ' + esc(d.voidedBy) + ': ' + esc(d.voidReason) : '') + '</td>' +
                '<td style="white-space:nowrap">' + esc(fmtKm(d.odometerKm)) + '</td>' +
                '<td>' + esc(d.createdBy) + '<br><span class="mv-muted">' + esc(fmtDateTime(d.createdAt)) + '</span></td>' +
                '<td>' + (d.voided || !isAdmin ? '' : '<button type="button" class="mv-btn-link" data-void="' + esc(r.id) + '">Anular</button>') + '</td>' +
                '</tr>';
        }).join('');
        $('#mv-history-empty').hidden = rows.length > 0;
        $('#mv-history-count').textContent = rows.length ? rows.length + (rows.length === 1 ? ' movimiento' : ' movimientos') : '';
    }

    function csvCell(v) {
        var s = String(v === undefined || v === null ? '' : v);
        if (/^[=+\-@]/.test(s)) { s = "'" + s; }
        return '"' + s.replace(/"/g, '""') + '"';
    }

    function exportCsv() {
        var rows = filteredHistory();
        var head = ['Fecha', 'Movimiento', 'Patente', 'Vehiculo Geotab', 'Chasis', 'Marca', 'Modelo', 'Serie GO', 'Arnes', 'Camara',
            'Serie camara', 'Odometro km', 'Detalle', 'Registrado por', 'Fecha de registro', 'Anulado', 'Motivo anulacion'];
        var lines = [head.map(csvCell).join(';')].concat(rows.map(function (r) {
            var d = r.details, c = d.camera || {};
            return [fmtDay(d.date), TYPE_LABEL[d.movementType], d.licensePlate, d.deviceName, d.chassis, d.vehicleBrand, d.vehicleModel,
                d.goSerial, d.harness, c.model, c.serial, d.odometerKm, detailOf(d), d.createdBy, fmtDateTime(d.createdAt),
                d.voided ? 'Si' : 'No', d.voidReason].map(csvCell).join(';');
        }));
        var blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
        var a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = 'movimientos_' + today() + '.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        setTimeout(function () { URL.revokeObjectURL(a.href); }, 1000);
    }

    // ------------------------------------------------------------------
    // Anular
    // ------------------------------------------------------------------
    function openVoid(id) {
        if (!isAdmin) { return; }
        voidTarget = records.filter(function (r) { return r.id === id; })[0] || null;
        if (!voidTarget) { return; }
        var d = voidTarget.details, f = $('#mv-void');
        f.reset();
        clearErrors(f);
        $('#mv-void-desc').textContent = TYPE_LABEL[d.movementType] + ' · ' + (d.licensePlate || d.deviceName) + ' · ' + fmtDay(d.date) +
            '. El registro queda en el historial como anulado.';
        $('#mv-modal').hidden = false;
        f.elements.reason.focus();
    }
    function closeVoid() {
        voidTarget = null;
        $('#mv-modal').hidden = true;
    }
    function submitVoid(e) {
        e.preventDefault();
        if (saving || !voidTarget || !isAdmin) { return; }
        var f = e.target, reason = f.elements.reason.value.trim();
        clearErrors(f);
        if (!reason) { fieldError(f, 'reason', 'Indicá el motivo.'); return; }
        var target = voidTarget, btn = f.querySelector('button[type=submit]');
        var details = {};
        Object.keys(target.details).forEach(function (k) { details[k] = target.details[k]; });
        details.voided = true;
        details.voidReason = reason;
        details.voidedBy = userName;
        details.voidedAt = new Date().toISOString();
        var entity = { id: target.id, addInId: ADDIN_ID, groups: target.groups || [{ id: 'GroupCompanyId' }], details: details };
        if (target.version) { entity.version = target.version; }
        saving = true;
        btn.disabled = true;
        call('Set', { typeName: 'AddInData', entity: entity }).then(function () {
            target.details = details;
            closeVoid();
            showMsg('Movimiento anulado.', 'ok');
            renderHistory();
            if (selected) { renderVehicle(); }
            return loadData();
        }).catch(function (err) {
            showMsg('No se pudo anular: ' + errText(err), 'err');
            closeVoid();
        }).then(function () {
            saving = false;
            btn.disabled = false;
        });
    }

    // ------------------------------------------------------------------
    // Eventos
    // ------------------------------------------------------------------
    function switchTab(name) {
        $$('.mv-tab').forEach(function (t) { t.classList.toggle('is-active', t.getAttribute('data-tab') === name); });
        ['install', 'change', 'history'].forEach(function (t) { $('#mv-tab-' + t).hidden = t !== name; });
        clearMsg();
        if (name === 'history') { renderHistory(); }
    }

    function bindEvents() {
        $$('.mv-tab').forEach(function (t) {
            t.addEventListener('click', function () { switchTab(t.getAttribute('data-tab')); });
        });
        $$('.mv-harness').forEach(function (s) { s.innerHTML = options(HARNESSES, 'Elegí el arnés'); });
        $$('.mv-camera').forEach(function (s) { s.innerHTML = options(CAMERAS, 'Sin cámara'); });

        var inst = $('#mv-install');
        inst.addEventListener('submit', submitInstall);
        $('#mv-install-clear').addEventListener('click', function () { resetInstall(); clearMsg(); });
        inst.elements.cameraModel.addEventListener('change', function () {
            syncCameraSerial(inst.elements.cameraModel, inst.elements.cameraSerial, 'Elegí primero marca y modelo');
        });

        $('#mv-search').addEventListener('submit', submitSearch);
        $('#mv-results').addEventListener('click', function (e) {
            var b = e.target.closest('[data-device]');
            if (!b) { return; }
            var id = b.getAttribute('data-device');
            selectDevice(devices.filter(function (d) { return d.id === id; })[0]);
        });
        // El formulario se vuelve a dibujar: se delega en el contenedor.
        var veh = $('#mv-vehicle');
        veh.addEventListener('change', syncChangeForm);
        veh.addEventListener('submit', function (e) { e.preventDefault(); submitChange(e.target); });

        ['#mv-h-text', '#mv-h-type', '#mv-h-from', '#mv-h-to', '#mv-h-void'].forEach(function (s) {
            $(s).addEventListener('input', renderHistory);
            $(s).addEventListener('change', renderHistory);
        });
        $('#mv-h-export').addEventListener('click', exportCsv);
        $('#mv-history').addEventListener('click', function (e) {
            var b = e.target.closest('[data-void]');
            if (b) { openVoid(b.getAttribute('data-void')); }
        });
        $('#mv-void').addEventListener('submit', submitVoid);
        $('#mv-void-cancel').addEventListener('click', closeVoid);
    }

    // ------------------------------------------------------------------
    // Ciclo de vida
    // ------------------------------------------------------------------
    return {
        initialize: function (freshApi, freshState, callback) {
            api = freshApi;
            root = document.getElementById('movimientos');
            if (!bound) { bindEvents(); bound = true; resetInstall(); }
            $('#mv-install').elements.date.max = today();
            var done = false;
            function finish() { if (!done) { done = true; callback(); } }
            try {
                api.getSession(function (s) {
                    userName = s && s.userName ? s.userName : '';
                    finish();
                });
            } catch (e) {
                finish();
            }
        },
        focus: function (freshApi) {
            api = freshApi;
            loadData();
        },
        blur: function () {
            closeVoid();
        }
    };
};

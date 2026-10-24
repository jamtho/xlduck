// Listen for messages from C#
        window.chrome.webview.addEventListener('message', event => {
            try {
                const state = JSON.parse(event.data);
                renderState(state);
            } catch (e) {
                console.error('Failed to parse state:', e);
            }
        });

        let renderGeneration = 0;
        let activeView = null;
        let activeMap = null;
        function disposeChart() {
            if (activeView) { activeView.finalize(); activeView = null; }
            if (activeMap) { activeMap.remove(); activeMap = null; }
        }
        function renderState(state) {
            const titleEl = document.getElementById('title');
            const handleEl = document.getElementById('handle');
            const contentEl = document.getElementById('content');

            let statusEl = document.getElementById('computation-status');
            if (!statusEl) {
                statusEl = document.createElement('div'); statusEl.id = 'computation-status';
                statusEl.style.cssText = 'padding:12px;color:#665000;background:#fff6d8;display:none';
                contentEl.before(statusEl);
            }
            if (state.kind === 'status') {
                statusEl.textContent = state.message || 'Computing'; statusEl.style.display = 'block';
                if (!state.retainPrevious) { renderGeneration++; disposeChart(); contentEl.replaceChildren(); }
                contentEl.style.opacity = state.retainPrevious ? '0.45' : '1';
                return;
            }
            renderGeneration++; disposeChart();
            statusEl.style.display = 'none'; contentEl.style.opacity = '1';

            titleEl.textContent = state.title || 'Preview';
            handleEl.textContent = state.handle || '';

            switch (state.kind) {
                case 'empty':
                    contentEl.innerHTML = `<div class="empty-state">${escapeHtml(state.message || 'Select a handle')}</div>`;
                    break;

                case 'error':
                    let errorHtml = `<div class="error-message">${escapeHtml(state.message || 'Error')}</div>`;
                    if (state.resolvedSql) {
                        errorHtml += '<div class="frag-section">';
                        errorHtml += '<div class="section-title">Resolved SQL</div>';
                        errorHtml += `<pre class="frag-sql">${escapeHtml(formatSql(state.resolvedSql))}</pre>`;
                        errorHtml += '</div>';
                    }
                    contentEl.innerHTML = errorHtml;
                    break;

                case 'table':
                    contentEl.innerHTML = renderTable(state.table);
                    break;

                case 'frag':
                    contentEl.innerHTML = renderFragment(state.frag);
                    break;

                case 'plot':
                    renderPlot(contentEl, state.plot);
                    break;

                default:
                    contentEl.innerHTML = `<div class="empty-state">Unknown state: ${escapeHtml(state.kind)}</div>`;
            }
        }

        function renderTable(table) {
            let html = '';

            // SQL section (if available)
            if (table.sql) {
                html += '<div class="frag-section">';
                html += '<div class="section-title">SQL</div>';
                html += `<pre class="frag-sql">${escapeHtml(formatSql(table.sql))}</pre>`;

                if (table.args && table.args.length > 0) {
                    html += '<div class="frag-args">';
                    html += '<div class="section-title">Arguments</div>';
                    for (const arg of table.args) {
                        html += '<div class="frag-arg">';
                        html += `<span class="frag-arg-name">${escapeHtml(arg.name)}</span>`;
                        html += `<span class="frag-arg-value">${escapeHtml(arg.value)}</span>`;
                        html += '</div>';
                    }
                    html += '</div>';
                }

                html += '</div>';
            }

            // Schema section
            html += '<div class="schema-section">';
            html += '<div class="section-title">Schema</div>';
            html += '<table class="schema-table">';
            html += '<tr><th>Column</th><th>Type</th></tr>';
            for (const col of table.columns) {
                html += `<tr><td>${escapeHtml(col.name)}</td><td class="type">${escapeHtml(col.type)}</td></tr>`;
            }
            html += '</table>';
            html += '</div>';

            // Data section
            html += '<div class="data-section">';
            html += `<div class="data-info">Showing ${table.previewRowCount} of ${table.rowCount.toLocaleString()} rows</div>`;
            html += '<div class="data-grid-container">';
            html += '<table class="data-grid">';

            // Header row
            html += '<tr>';
            for (const col of table.columns) {
                html += `<th>${escapeHtml(col.name)}</th>`;
            }
            html += '</tr>';

            // Data rows
            for (const row of table.rows) {
                html += '<tr>';
                for (const cell of row) {
                    if (cell === null) {
                        html += '<td class="null">NULL</td>';
                    } else {
                        html += `<td title="${escapeHtml(String(cell))}">${escapeHtml(formatCell(cell))}</td>`;
                    }
                }
                html += '</tr>';
            }

            html += '</table>';
            html += '</div>';
            html += '</div>';

            return html;
        }

        function renderFragment(frag) {
            let html = '';

            html += '<div class="frag-section">';
            html += '<div class="section-title">SQL</div>';
            html += `<pre class="frag-sql">${escapeHtml(formatSql(frag.sql))}</pre>`;

            if (frag.args && frag.args.length > 0) {
                html += '<div class="frag-args">';
                html += '<div class="section-title">Parameters</div>';
                for (const arg of frag.args) {
                    html += '<div class="frag-arg">';
                    html += `<span class="frag-arg-name">${escapeHtml(arg.name)}</span>`;
                    html += `<span class="frag-arg-value">${escapeHtml(arg.value)}</span>`;
                    html += '</div>';
                }
                html += '</div>';
            }

            html += '</div>';

            return html;
        }

        function formatCell(value) {
            if (typeof value === 'string' && value.length > 50) {
                return value.substring(0, 47) + '...';
            }
            return String(value);
        }

        function formatSql(sql) {
            if (typeof sqlFormatter !== 'undefined') {
                try {
                    sql = sqlFormatter.format(sql, { language: 'duckdb', tabWidth: 2 });
                } catch (e) {
                    // Fall through with unformatted sql
                }
            }
            return sql;
        }

        function escapeHtml(text) {
            return String(text ?? '').replace(/[&<>"']/g, character => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[character]));
        }

        function renderPlot(container, plotData) {
            const generation = renderGeneration;
            // Check for error
            if (plotData.error) {
                container.innerHTML = `<div class="plot-section"><div class="plot-error">${escapeHtml(plotData.error)}</div></div>`;
                return;
            }

            // Build HTML structure
            let html = '<div class="plot-section">';
            html += '<div id="plot-target" class="plot-container"></div>';
            html += '<div class="plot-info">';
            html += `<div class="plot-info-row"><span class="plot-info-label">Template:</span><span class="plot-info-value">${escapeHtml(plotData.template)}</span></div>`;
            html += `<div class="plot-info-row"><span class="plot-info-label">Rows:</span><span class="plot-info-value">${plotData.rowCount.toLocaleString()}${plotData.isSample ? ' — chart uses a sample of ' + plotData.rows.length.toLocaleString() : ''}</span></div>`;

            // Show overrides
            for (const [key, value] of Object.entries(plotData.overrides || {})) {
                html += `<div class="plot-info-row"><span class="plot-info-label">${escapeHtml(key)}:</span><span class="plot-info-value">${escapeHtml(value)}</span></div>`;
            }

            html += '</div>';
            html += '</div>';
            container.innerHTML = html;

            // Build data values array
            const dataValues = [];
            try {
            for (let rowIndex = 0; rowIndex < plotData.rows.length; rowIndex++) {
                const row = plotData.rows[rowIndex];
                const obj = Object.create(null);
                for (let i = 0; i < plotData.columns.length; i++) {
                    const col = plotData.columns[i];
                    const val = chartValue(row[i], plotData.kinds?.[rowIndex]?.[i], plotData.types[i]);
                    if (isIntervalType(plotData.types[i])) obj[durationField(plotData.columns, i)] = formatDuration(val);
                    obj[col] = val;
                }
                dataValues.push(obj);
            }
            } catch (error) {
                container.querySelector('#plot-target').textContent = error.message;
                return;
            }

            // Branch: map template uses Leaflet, everything else uses Vega-Lite
            if (plotData.template === 'map') {
                requestAnimationFrame(() => {
                    if (generation !== renderGeneration) return;
                    renderMap(document.getElementById('plot-target'), plotData, dataValues);
                });
                return;
            }

            // Build Vega-Lite spec
            let spec;
            try { spec = buildVegaSpec(plotData.template, plotData.overrides, plotData.columns, plotData.types, dataValues); }
            catch (error) { container.querySelector('#plot-target').textContent = error.message; return; }

            // Render with vegaEmbed (defer to next frame so container has layout dimensions)
            requestAnimationFrame(() => {
                if (generation !== renderGeneration) return;
                if (typeof vegaEmbed !== 'undefined') {
                    try {
                        const target = document.getElementById('plot-target');
                        vegaEmbed(target, spec, {
                            actions: false,
                            ast: true,
                            expr: vega.expressionInterpreter,
                            renderer: 'canvas'
                        }).then(result => {
                            if (generation !== renderGeneration) result.view.finalize();
                            else activeView = result.view;
                        }).catch(err => {
                            if (generation !== renderGeneration) return;
                            target.innerHTML =
                                `<div class="plot-error">Render error: ${escapeHtml(err.message)}</div>`;
                        });
                    } catch (err) {
                        document.getElementById('plot-target').innerHTML =
                            `<div class="plot-error">Spec error: ${escapeHtml(err.message)}</div>`;
                    }
                } else {
                    document.getElementById('plot-target').innerHTML =
                        '<div class="plot-error">Chart libraries are missing. Repair the XlDuck installation.</div>';
                }
            });
        }

        function validateMapData(plotData, dataValues) {
            const { lat, lon, color, basemap } = plotData.overrides;
            if (!lat || !lon) throw new Error('Map requires lat and lon column names.');
            for (const name of [lat, lon, color].filter(Boolean)) {
                if (!plotData.columns.includes(name)) throw new Error('Unknown map column: ' + name);
            }
            for (const name of [lat, lon]) {
                const type = plotData.types[plotData.columns.indexOf(name)];
                if (!isQuantitative(type) || isIntervalType(type))
                    throw new Error('Map coordinates must be numeric. Cast latitude and longitude explicitly in SQL.');
            }
            if (basemap && !['osm', 'satellite'].includes(basemap.toLowerCase()))
                throw new Error('Map basemap must be osm or satellite.');
            for (const row of dataValues) {
                if (!Number.isFinite(row[lat]) || !Number.isFinite(row[lon]) ||
                    Math.abs(row[lat]) > 90 || Math.abs(row[lon]) > 180)
                    throw new Error('Map coordinates must be finite numbers within latitude -90 to 90 and longitude -180 to 180. Filter NULL or invalid coordinates in SQL.');
            }
        }

        function renderMap(target, plotData, dataValues) {
            try { validateMapData(plotData, dataValues); }
            catch (error) {
                target.innerHTML = `<div class="plot-error">${escapeHtml(error.message)}</div>`;
                return;
            }
            if (typeof L === 'undefined') {
                target.innerHTML = '<div class="plot-error">Map library is missing. Repair the XlDuck installation.</div>';
                return;
            }

            const latField = plotData.overrides.lat;
            const lonField = plotData.overrides.lon;
            const colorField = plotData.overrides.color;
            const title = plotData.overrides.title;

            // Set up the map container with canvas renderer for performance
            target.classList.add('map-container');
            const map = L.map(target, { preferCanvas: true });
            activeMap = map;

            const layer = (plotData.overrides.basemap || 'osm').toLowerCase();
            if (layer === 'satellite') {
                L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
                    attribution: 'Tiles &copy; Esri',
                    maxZoom: 19
                }).addTo(map);
            } else {
                L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
                    attribution: '&copy; OpenStreetMap contributors',
                    maxZoom: 19
                }).addTo(map);
            }

            // Color palette for groups
            const palette = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf'];

            // Build popup content for a data row
            function popupContent(row) {
                let html = '<table style="font-size:11px;border-collapse:collapse">';
                for (const [key, val] of Object.entries(row)) {
                    html += `<tr><td style="font-weight:600;padding:1px 6px 1px 0">${escapeHtml(key)}</td><td>${escapeHtml(String(val ?? ''))}</td></tr>`;
                }
                html += '</table>';
                return html;
            }

            const allPoints = [];
            const layers = L.layerGroup();

            if (colorField) {
                // Assign a color per distinct value
                const colorMap = Object.create(null);
                let colorIdx = 0;
                for (const row of dataValues) {
                    const key = String(row[colorField] ?? '');
                    if (!(key in colorMap)) {
                        colorMap[key] = palette[colorIdx % palette.length];
                        colorIdx++;
                    }
                }

                // Count group sizes for z-ordering (largest first, smallest on top)
                const groupSizes = Object.create(null);
                for (const row of dataValues) {
                    const key = String(row[colorField] ?? '');
                    groupSizes[key] = (groupSizes[key] || 0) + 1;
                }

                // Draw per-segment polylines in data order: split when color changes
                let segPoints = [];
                let segColor = null;
                for (const row of dataValues) {
                    const lat = row[latField];
                    const lon = row[lonField];
                    const pt = [lat, lon];
                    allPoints.push(pt);
                    const color = colorMap[String(row[colorField] ?? '')];

                    if (segColor !== null && color !== segColor) {
                        // Color changed — flush segment (include last point for continuity)
                        if (segPoints.length > 1) {
                            layers.addLayer(L.polyline(segPoints, { color: segColor, weight: 2, opacity: 0.6, interactive: false }));
                        }
                        segPoints = [segPoints[segPoints.length - 1]]; // bridge point
                    }
                    segColor = color;
                    segPoints.push(pt);
                }
                if (segPoints.length > 1) {
                    layers.addLayer(L.polyline(segPoints, { color: segColor, weight: 2, opacity: 0.6, interactive: false }));
                }

                // Draw circle markers sorted by group size desc (largest/most common first,
                // smallest/rarest on top so they remain visible)
                const sortedRows = [...dataValues].sort((a, b) => {
                    const sA = groupSizes[String(a[colorField] ?? '')] || 0;
                    const sB = groupSizes[String(b[colorField] ?? '')] || 0;
                    return sB - sA;
                });
                for (const row of sortedRows) {
                    const lat = row[latField];
                    const lon = row[lonField];
                    const pt = [lat, lon];
                    const color = colorMap[String(row[colorField] ?? '')];
                    layers.addLayer(L.circleMarker(pt, {
                        radius: 4, color: color, fillColor: color, fillOpacity: 0.7, weight: 1
                    }).bindPopup(popupContent(row)));
                }
            } else {
                // Single group
                const points = [];
                const color = palette[0];

                for (const row of dataValues) {
                    const lat = row[latField];
                    const lon = row[lonField];
                    const pt = [lat, lon];
                    points.push(pt);
                    allPoints.push(pt);

                    layers.addLayer(L.circleMarker(pt, {
                        radius: 4, color: color, fillColor: color, fillOpacity: 0.7, weight: 1
                    }).bindPopup(popupContent(row)));
                }

                if (points.length > 1) {
                    layers.addLayer(L.polyline(points, { color: color, weight: 2, opacity: 0.6, interactive: false }));
                }
            }

            layers.addTo(map);

            // Fit bounds to data extent
            if (allPoints.length > 0) {
                map.fitBounds(allPoints, { padding: [20, 20] });
            } else {
                map.setView([0, 0], 2);
            }

            if (title) {
                const titleControl = L.control({ position: 'topleft' });
                titleControl.onAdd = function() {
                    const div = L.DomUtil.create('div', '');
                    div.style.cssText = 'background:white;padding:4px 8px;border-radius:4px;font-weight:600;font-size:13px;box-shadow:0 1px 3px rgba(0,0,0,.3)';
                    div.textContent = title;
                    return div;
                };
                titleControl.addTo(map);
            }
        }

        function buildVegaSpec(template, overrides, columns, types, dataValues) {
            const required = template === 'histogram' ? ['x'] : template === 'heatmap' ? ['x','y','value'] : ['x','y'];
            for (const option of required)
                if (!overrides[option] || !columns.includes(overrides[option])) throw new Error(`Choose a valid ${option} column for this chart.`);
            if (template === 'histogram' && !isQuantitative(types[columns.indexOf(overrides.x)]) && !isIntervalType(types[columns.indexOf(overrides.x)]))
                throw new Error('Histograms need numeric values. Convert the source column explicitly in SQL.');
            const xField = overrides.x;
            const yField = overrides.y;
            const colorField = overrides.color;
            const valueField = overrides.value;
            const labelField = overrides.label;
            const tooltipField = overrides.tooltip;
            const title = overrides.title;

            // Get types for fields
            const xType = getVegaType(xField, columns, types);

            let spec;

            // Handle special templates
            if (template === 'histogram') {
                spec = {
                    "$schema": "https://vega.github.io/schema/vega-lite/v5.json",
                    "width": "container",
                    "height": 300,
                    "data": { "values": dataValues },
                    "mark": { "type": "bar", "tooltip": {"content": "data"} },
                    "encoding": {
                        "x": { "field": xField, "bin": true, "type": xType === "nominal" ? "quantitative" : xType },
                        "y": { "aggregate": "count" }
                    }
                };
            } else if (template === 'heatmap') {
                const yType = getVegaType(yField, columns, types);
                spec = {
                    "$schema": "https://vega.github.io/schema/vega-lite/v5.json",
                    "width": "container",
                    "height": 300,
                    "data": { "values": dataValues },
                    "mark": { "type": "rect", "tooltip": {"content": "data"} },
                    "encoding": {
                        "x": { "field": xField, "type": xType === "quantitative" ? "ordinal" : xType },
                        "y": { "field": yField, "type": yType === "quantitative" ? "ordinal" : yType },
                        "color": { "field": valueField, "type": "quantitative", "aggregate": "mean" }
                    }
                };
            } else if (template === 'boxplot') {
                spec = {
                    "$schema": "https://vega.github.io/schema/vega-lite/v5.json",
                    "width": "container",
                    "height": 300,
                    "data": { "values": dataValues },
                    "mark": { "type": "boxplot", "extent": "min-max", "tooltip": {"content": "data"} },
                    "encoding": {
                        "x": { "field": xField, "type": xType === "quantitative" ? "nominal" : xType },
                        "y": { "field": yField, "type": "quantitative" }
                    }
                };
            } else {
                // Standard templates: bar, line, point, area
                const yType = getVegaType(yField, columns, types);
                spec = {
                    "$schema": "https://vega.github.io/schema/vega-lite/v5.json",
                    "width": "container",
                    "height": 300,
                    "data": { "values": dataValues },
                    "mark": {
                        "type": template,
                        "tooltip": {"content": "data"}
                    },
                    "encoding": {
                        "x": { "field": xField, "type": xType },
                        "y": { "field": yField, "type": yType }
                    }
                };

                // Line/area: draw in data order instead of sorting by x
                if (template === 'line' || template === 'area') {
                    spec.encoding.order = false;
                }

                // Add point markers for lines
                if (template === 'line') {
                    spec.mark.point = true;
                }

                // Fit quantitative axes to data range (skip bar which needs zero baseline)
                if (template !== 'bar') {
                    if (xType === 'quantitative') {
                        spec.encoding.x.scale = { "zero": false };
                    }
                    if (yType === 'quantitative') {
                        spec.encoding.y.scale = { "zero": false };
                    }
                }

                // Add color encoding if specified
                if (colorField) {
                    const colorType = getVegaType(colorField, columns, types);
                    spec.encoding.color = { "field": colorField, "type": colorType };
                }
            }

            // Tooltip: specific field overrides the mark-level "all data" default
            if (tooltipField && spec.encoding) {
                spec.encoding.tooltip = { "field": tooltipField, "type": getVegaType(tooltipField, columns, types) };
            }

            // Add title if specified
            if (title) {
                spec.title = title;
            }

            // Apply explicit scale domains if provided (merge with existing scale)
            if (spec.encoding && spec.encoding.x) {
                const xDomain = parseDomain(overrides.xmin, overrides.xmax);
                if (xDomain) {
                    spec.encoding.x.scale = Object.assign({}, spec.encoding.x.scale, { "domain": xDomain });
                }
            }
            if (spec.encoding && spec.encoding.y) {
                const yDomain = parseDomain(overrides.ymin, overrides.ymax);
                if (yDomain) {
                    spec.encoding.y.scale = Object.assign({}, spec.encoding.y.scale, { "domain": yDomain });
                }
            }

            // Fit the entire chart (including axes/labels) within the container
            spec.autosize = { "type": "fit", "contains": "padding" };

            // Format INTERVAL axes as H:MM:SS instead of raw seconds
            const durationLabelExpr =
                "(datum.value < 0 ? '-' : '') + floor(abs(datum.value) / 3600) + ':' + " +
                "pad('' + floor((abs(datum.value) % 3600) / 60), 2, '0') + ':' + " +
                "pad('' + floor(abs(datum.value) % 60), 2, '0')";
            if (spec.encoding) {
                if (xField && isIntervalType((types[columns.indexOf(xField)] || ''))) {
                    spec.encoding.x.axis = Object.assign({}, spec.encoding.x.axis, { "labelExpr": durationLabelExpr });
                }
                if (yField && isIntervalType((types[columns.indexOf(yField)] || ''))) {
                    spec.encoding.y.axis = Object.assign({}, spec.encoding.y.axis, { "labelExpr": durationLabelExpr });
                }

                // Show formatted H:MM:SS in tooltips instead of raw seconds
                if (!tooltipField && columns.some((c, i) => isIntervalType(types[i]))) {
                    spec.encoding.tooltip = columns.map((col, i) => {
                        if (isIntervalType(types[i])) {
                            return {"field": durationField(columns, i), "type": "nominal", "title": col};
                        }
                        return {"field": col, "type": getVegaType(col, columns, types)};
                    });
                    if (spec.mark && spec.mark.tooltip) {
                        delete spec.mark.tooltip;
                    }
                }
            }

            // Enable interactive pan and zoom (drag to pan, scroll to zoom)
            const zoomParams = [{ "name": "grid", "select": "interval", "bind": "scales" }];

            // Add text labels as a layer if specified
            if (labelField) {
                const baseMark = spec.mark;
                const baseEncoding = spec.encoding;
                delete spec.mark;
                delete spec.encoding;

                // Text layer gets only x, y positioning and the label text
                const textEncoding = {
                    "x": baseEncoding.x,
                    "y": baseEncoding.y,
                    "text": { "field": labelField, "type": "nominal" }
                };

                spec.layer = [
                    { "mark": baseMark, "encoding": baseEncoding, "params": zoomParams },
                    {
                        "mark": { "type": "text", "dy": -10, "fontSize": 10 },
                        "encoding": textEncoding
                    }
                ];
            } else {
                spec.params = zoomParams;
            }

            // SQL column names are literal fields. Vega's field syntax otherwise treats
            // dots/brackets as nested access and silently plots missing values.
            const escapeFields = value => {
                if (Array.isArray(value)) { value.forEach(escapeFields); return; }
                if (!value || typeof value !== 'object') return;
                for (const [key, child] of Object.entries(value)) {
                    if (key === 'field' && typeof child === 'string') value[key] = child.replace(/[\\.\[\]]/g, '\\$&');
                    else escapeFields(child);
                }
            };
            if (spec.encoding) escapeFields(spec.encoding);
            if (spec.layer) for (const layer of spec.layer) escapeFields(layer.encoding);
            return spec;
        }

        function durationField(columns, index) {
            let field = '__xlduck_duration_' + index;
            while (columns.includes(field)) field += '_';
            return field;
        }

        function parseDomain(minVal, maxVal) {
            if (minVal == null && maxVal == null) return null;
            const min = minVal != null ? parseFloat(minVal) : undefined;
            const max = maxVal != null ? parseFloat(maxVal) : undefined;
            if (min != null && !isNaN(min) && max != null && !isNaN(max)) return [min, max];
            if (min != null && !isNaN(min)) return [min, undefined];
            if (max != null && !isNaN(max)) return [undefined, max];
            return null;
        }

        // Format total seconds as H:MM:SS (matches axis labels)
        function formatDuration(secs) {
            if (secs == null) return null;
            const sign = secs < 0 ? '-' : '';
            const abs = Math.abs(secs);
            const h = Math.floor(abs / 3600);
            const m = String(Math.floor((abs % 3600) / 60)).padStart(2, '0');
            const s = String(Math.floor(abs % 60)).padStart(2, '0');
            return sign + h + ':' + m + ':' + s;
        }

        function getVegaType(fieldName, columns, types) {
            const idx = columns.indexOf(fieldName);
            if (idx < 0) return 'nominal';

            const duckType = (types[idx] || '').toUpperCase();

            // INTERVAL is duration data - plot as quantitative (seconds), not temporal
            if (duckType === 'INTERVAL') {
                return 'quantitative';
            }
            if (/^(DATE|TIME(?: WITH TIME ZONE)?|TIMESTAMP(?:_S|_MS|_NS)?(?: WITH TIME ZONE)?)$/.test(duckType)) {
                return 'temporal';
            }
            if (isQuantitative(duckType)) {
                return 'quantitative';
            }
            return 'nominal';
        }

        function isQuantitative(duckType) {
            const t = (duckType || '').toUpperCase();
            return /^(U?(?:TINYINT|SMALLINT|INTEGER|BIGINT|HUGEINT)|INT|FLOAT|DOUBLE|REAL|(?:DECIMAL|NUMERIC)(?:\(\d+,\d+\))?)$/.test(t);
        }

        function chartValue(value, kind, type) {
            if (value == null || kind === 'Null' || kind === 'Empty') return null;
            if (kind === 'Interval' || isIntervalType(type)) {
                const seconds = parseDurationToSeconds(value);
                if (seconds == null) throw new Error('This interval cannot be represented as fixed seconds. Convert it explicitly in SQL.');
                return seconds;
            }
            if (kind === 'Text' || kind === 'Json') return String(value);
            if (kind === 'Boolean') return value === true || value === 'true';
            if (kind === 'Decimal' || /^(DECIMAL|NUMERIC)/i.test(type || ''))
                throw new Error('Exact decimals need an explicit CAST AS DOUBLE for charting. Table previews preserve their exact values.');
            if (kind === 'Integer' || /^(U?(?:TINYINT|SMALLINT|INTEGER|BIGINT|HUGEINT)|INT)$/i.test(type || '')) {
                const integer = BigInt(value);
                if (integer > 9007199254740991n || integer < -9007199254740991n)
                    throw new Error('This integer exceeds chart precision. Cast it to VARCHAR for a label, or DOUBLE to choose approximation.');
                return Number(integer);
            }
            if (kind === 'Number' || isQuantitative(type)) {
                const number = Number(value);
                if (!Number.isFinite(number)) throw new Error('Chart data contains an invalid number.');
                return number;
            }
            return value;
        }

        function isIntervalType(duckType) {
            return (duckType || '').toUpperCase() === 'INTERVAL';
        }

        // Parse .NET TimeSpan format [-][d.]hh:mm:ss[.fffffff] to total seconds
        function parseDurationToSeconds(str) {
            if (str == null) return null;
            const s = String(str).trim();
            if (!s) return null;
            const duck = s.match(/^([+-]?\d+) days?\s+([+-]?)(\d+):(\d+):(\d+(?:\.\d+)?)$/);
            if (duck) return Number(duck[1]) * 86400 + (duck[2] === '-' ? -1 : 1) * (Number(duck[3]) * 3600 + Number(duck[4]) * 60 + Number(duck[5]));
            const m = s.match(/^(-)?(?:(\d+)\.)?(\d+):(\d+):(\d+)(?:\.(\d+))?$/);
            if (!m) return null;
            const sign = m[1] ? -1 : 1;
            const days = m[2] ? parseInt(m[2], 10) : 0;
            const hours = parseInt(m[3], 10);
            const mins = parseInt(m[4], 10);
            const secs = parseInt(m[5], 10) + (m[6] ? Number('0.' + m[6]) : 0);
            return sign * (days * 86400 + hours * 3600 + mins * 60 + secs);
        }

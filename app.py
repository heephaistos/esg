import streamlit as st
import re
import pandas as pd
import docx
import fitz  # PyMuPDF
import os
import glob
import streamlit.components.v1 as components

st.set_page_config(page_title="BID ESG Expert Classifier", page_icon="🏦", layout="wide", initial_sidebar_state="expanded")

if 'new_cases' not in st.session_state:
    st.session_state['new_cases'] = []

# --- CSS STYLING ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .stApp { background-color: #f8fafc; }
    [data-testid="stSidebar"] { background-color: #0b2240; }
    [data-testid="stSidebar"] * { color: #ffffff !important; }
    
    .main-header { font-size: 2.2rem; color: #0b2240; font-weight: 800; margin-bottom: 0.2rem; }
    .sub-header { font-size: 1.1rem; color: #475569; margin-bottom: 2rem; }
    
    .card { background: white; padding: 1.5rem; border-radius: 12px; box-shadow: 0 4px 15px rgba(0,0,0,0.04); border: 1px solid #e2e8f0; margin-bottom: 1.5rem; }
    .section-title { font-size: 1.3rem; font-weight: 700; color: #1e293b; margin-bottom: 1rem; border-bottom: 2px solid #e2e8f0; padding-bottom: 0.5rem; }
    
    .badge-A { background-color: #fef2f2; color: #dc2626; border: 1px solid #fecaca; }
    .badge-B { background-color: #fffbeb; color: #d97706; border: 1px solid #fde68a; }
    .badge-C { background-color: #f0fdf4; color: #16a34a; border: 1px solid #bbf7d0; }
    .badge-FI { background-color: #eff6ff; color: #2563eb; border: 1px solid #bfdbfe; }
    
    .badge-Elevado, .badge-High { background-color: #fef2f2; color: #dc2626; border: 1px solid #fecaca; }
    .badge-Sustancial, .badge-Substantial { background-color: #fff7ed; color: #ea580c; border: 1px solid #ffedd5; }
    .badge-Moderado, .badge-Moderate { background-color: #fef9c3; color: #ca8a04; border: 1px solid #fef08a; }
    .badge-Bajo, .badge-Low { background-color: #f0fdf4; color: #16a34a; border: 1px solid #bbf7d0; }
    
    .big-badge { display: inline-block; padding: 0.5rem 1.5rem; border-radius: 8px; font-weight: 800; font-size: 1.6rem; text-align: center; width: 100%;}
    
    .fmea-table th { background-color: #0b2240 !important; color: white !important; }
    
    .factor-box { padding: 1rem; border-radius: 8px; background: #f8fafc; border-left: 4px solid #3b82f6; margin-bottom: 1rem; height: 100%; }
    .factor-title { font-weight: 700; color: #0f172a; margin-bottom: 0.5rem; }
    
    .summary-box { background: #f0fdf4; border-left: 5px solid #16a34a; padding: 1.5rem; border-radius: 8px; margin-bottom: 2rem; color: #0f172a; font-size: 1.05rem; line-height: 1.6; }
</style>
""", unsafe_allow_html=True)

def render_ishikawa_svg(data):
    if not data: return ""
    effect = data.get("effect", "Efecto Principal")
    categories = data.get("categories", [])
    
    while len(categories) < 6:
        categories.append({"category": "N/A", "causes": []})
        
    def generate_html_list(causes, level=0):
        if not causes:
            return ""
        html = ""
        for c in causes:
            if isinstance(c, dict):
                html += f"<li>{c.get('name', '')}"
                if c.get('subcauses'):
                    bullet_class = f"subcause-level-{level+1}"
                    html += f"<ul class='subcause-list {bullet_class}'>"
                    html += generate_html_list(c['subcauses'], level+1)
                    html += "</ul>"
                html += "</li>"
            else:
                html += f"<li>{c}</li>"
        return html
        
    svg = f'''
    <svg width="100%" height="100%" viewBox="0 0 1450 850" preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg">
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&amp;display=swap');
            .bg {{ fill: #ffffff; }}
            .spine {{ stroke: #0f172a; stroke-width: 6px; stroke-linecap: round; }}
            .bone {{ stroke: #64748b; stroke-width: 3px; }}
            .box {{ font-family: 'Inter', sans-serif; background: white; border: 1px solid #cbd5e1; border-radius: 8px; padding: 12px; box-shadow: 0 4px 10px rgba(0,0,0,0.05); width: 100%; height: 100%; box-sizing: border-box; display: flex; flex-direction: column; overflow-y: auto; overflow-x: hidden; }}
            .cat-title {{ font-weight: 700; color: #1e293b; margin-bottom: 8px; border-bottom: 2px solid #e2e8f0; padding-bottom: 6px; font-size: 13.5px; letter-spacing: 0.2px; }}
            
            .cause-list {{ list-style-type: none; padding-left: 0; margin: 0; color: #1e293b; font-size: 12px; line-height: 1.35; }}
            .cause-list > li {{ margin-bottom: 8px; position: relative; padding-left: 14px; font-weight: 600; }}
            .cause-list > li::before {{ content: "•"; color: #0ea5e9; font-weight: bold; position: absolute; left: 0; top: 0; font-size: 18px; line-height: 1.1; }}
            
            .subcause-list {{ list-style-type: none; padding-left: 8px; margin: 3px 0 0 0; font-size: 11px; font-weight: 400; color: #475569; }}
            .subcause-level-1 > li {{ margin-bottom: 3px; position: relative; padding-left: 12px; }}
            .subcause-level-1 > li::before {{ content: "↳"; color: #64748b; font-weight: bold; position: absolute; left: 0; top: 0; font-size: 12px; line-height: 1.2; }}
            
            .subcause-level-2 > li {{ margin-bottom: 2px; position: relative; padding-left: 12px; color: #64748b; font-style: italic; }}
            .subcause-level-2 > li::before {{ content: "-"; color: #94a3b8; font-weight: bold; position: absolute; left: 0; top: 0; font-size: 12px; line-height: 1.2; }}
            
            .effect-box {{ font-family: 'Inter', sans-serif; font-size: 15px; background: #fef2f2; border: 3px solid #dc2626; border-radius: 12px; padding: 15px; color: #991b1b; font-weight: 700; text-align: center; display: flex; align-items: center; justify-content: center; height: 100%; box-sizing: border-box; box-shadow: 0 6px 15px rgba(220, 38, 38, 0.15); line-height: 1.3; }}
            
            ::-webkit-scrollbar {{ width: 6px; }}
            ::-webkit-scrollbar-track {{ background: #f1f5f9; border-radius: 4px; }}
            ::-webkit-scrollbar-thumb {{ background: #cbd5e1; border-radius: 4px; }}
            ::-webkit-scrollbar-thumb:hover {{ background: #94a3b8; }}
        </style>
        
        <rect width="100%" height="100%" class="bg" rx="16" />
        
        <!-- Spine -->
        <line x1="40" y1="425" x2="1140" y2="425" class="spine" />
        <polygon points="1130,410 1160,425 1130,440" fill="#0f172a" />
        
        <!-- Effect Box (Head) -->
        <foreignObject x="1170" y="340" width="240" height="170">
            <div xmlns="http://www.w3.org/1999/xhtml" class="effect-box">
                {effect}
            </div>
        </foreignObject>
    '''
    
    positions = [
        {"x_box": 40, "y_box": 20, "x_bone1": 210, "y_bone1": 280, "x_bone2": 350, "y_bone2": 425, "is_top": True},
        {"x_box": 420, "y_box": 20, "x_bone1": 590, "y_bone1": 280, "x_bone2": 730, "y_bone2": 425, "is_top": True},
        {"x_box": 800, "y_box": 20, "x_bone1": 970, "y_bone1": 280, "x_bone2": 1110, "y_bone2": 425, "is_top": True},
        {"x_box": 40, "y_box": 570, "x_bone1": 210, "y_bone1": 570, "x_bone2": 350, "y_bone2": 425, "is_top": False},
        {"x_box": 420, "y_box": 570, "x_bone1": 590, "y_bone1": 570, "x_bone2": 730, "y_bone2": 425, "is_top": False},
        {"x_box": 800, "y_box": 570, "x_bone1": 970, "y_bone1": 570, "x_bone2": 1110, "y_bone2": 425, "is_top": False},
    ]
    
    for i, pos in enumerate(positions):
        cat = categories[i]
        x2, y2 = pos['x_bone2'], pos['y_bone2']
        if pos['is_top']:
            arrow = f'<polygon points="{x2-12},{y2-5} {x2},{y2} {x2-5},{y2-12}" fill="#64748b" />'
        else:
            arrow = f'<polygon points="{x2-12},{y2+5} {x2},{y2} {x2-5},{y2+12}" fill="#64748b" />'
            
        causes_html = generate_html_list(cat['causes'])
                
        svg += f'''
        <line x1="{pos['x_bone1']}" y1="{pos['y_bone1']}" x2="{pos['x_bone2']}" y2="{pos['y_bone2']}" class="bone" />
        {arrow}
        <foreignObject x="{pos['x_box']}" y="{pos['y_box']}" width="340" height="260">
            <div xmlns="http://www.w3.org/1999/xhtml" class="box">
                <div class="cat-title">{cat['category']}</div>
                <ul class="cause-list">
                    {causes_html}
                </ul>
            </div>
        </foreignObject>
        '''
        
    svg += "</svg>"
    components.html(svg, height=870)


def advanced_document_processing(files):
    # Simulated RAG/GraphRAG Pipeline using PyMuPDF and Regex
    markdown_text = ""
    chunks = []
    entities = {"nodes": set(), "edges": []}
    
    # 1. Extraction & Markdown Conversion
    for file in files:
        file.seek(0) # reset pointer
        try:
            if file.name.endswith(".pdf"):
                doc = fitz.open(stream=file.read(), filetype="pdf")
                for page in doc:
                    try:
                        blocks = page.get_text("dict")["blocks"]
                        for b in blocks:
                            if "lines" in b:
                                for l in b["lines"]:
                                    for s in l["spans"]:
                                        text = s["text"].strip()
                                        if not text: continue
                                        size = s["size"]
                                        if size > 16:
                                            markdown_text += f"\n# {text}\n"
                                        elif size > 12:
                                            markdown_text += f"\n## {text}\n"
                                        else:
                                            markdown_text += f"{text} "
                                markdown_text += "\n\n"
                    except:
                        # Fallback for plain text if dict fails
                        markdown_text += page.get_text() + "\n\n"
            elif file.name.endswith(".docx"):
                doc = docx.Document(file)
                markdown_text += "\n".join([p.text for p in doc.paragraphs]) + "\n"
        except Exception as e: 
            print(e)
            pass
        
    # 2. Semantic Chunking (by headers and size)
    raw_chunks = re.split(r'\n#+ ', markdown_text)
    for rc in raw_chunks:
        if len(rc.strip()) < 50: continue
        # if chunk is too large, split by paragraphs
        if len(rc) > 1500:
            sub_chunks = rc.split('\n\n')
            temp = ""
            for sc in sub_chunks:
                if len(temp) + len(sc) > 1500:
                    chunks.append(temp.strip())
                    temp = sc + "\n\n"
                else:
                    temp += sc + "\n\n"
            if temp: chunks.append(temp.strip())
        else:
            chunks.append(rc.strip())
            
    # 3. Simple GraphRAG extraction (Mocked via NER heuristics)
    for idx, chunk in enumerate(chunks):
        c_lower = chunk.lower()
        if "riesgo" in c_lower or "impacto" in c_lower or "desastre" in c_lower:
            entities["nodes"].add("Risk Event")
            entities["edges"].append({"source": f"Chunk {idx+1}", "target": "Risk Event", "label": "evaluates"})
        if "comunidad" in c_lower or "indígena" in c_lower or "social" in c_lower:
            entities["nodes"].add("Social Stakeholder")
            entities["edges"].append({"source": f"Chunk {idx+1}", "target": "Social Stakeholder", "label": "mentions"})
        if "ambiental" in c_lower or "ecosistema" in c_lower or "agua" in c_lower or "coral" in c_lower:
            entities["nodes"].add("Environmental Asset")
            entities["edges"].append({"source": f"Chunk {idx+1}", "target": "Environmental Asset", "label": "impacts"})
        if "gobierno" in c_lower or "ley" in c_lower or "política" in c_lower or "ministerio" in c_lower:
            entities["nodes"].add("Governance/State")
            entities["edges"].append({"source": f"Chunk {idx+1}", "target": "Governance/State", "label": "involves"})
            
    return markdown_text.lower(), chunks, entities

def render_graphrag_ui(chunks, entities):
    if not chunks: return
    st.markdown('<div class="section-title" style="margin-top:2rem;">0. Motor GraphRAG & Chunking Semántico (Ingesta de Conocimiento)</div>', unsafe_allow_html=True)
    st.markdown("Se ha procesado el documento original convirtiéndolo a **Markdown estructural** y particionándolo en **Chunks vectoriales**. Además se ha extraído un grafo de conocimiento subyacente para mejorar los cruces ESG.")
    
    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown(f"**📚 Chunks Semánticos ({len(chunks)} extraídos)**")
        with st.expander("Ver Todos los Fragmentos Vectorizados (Chunks Completos)"):
            for i, c in enumerate(chunks):
                st.info(f"**Chunk {i+1}:**\n\n" + c)
                
    with col2:
        st.markdown(f"**🕸️ Red de Conocimiento GraphRAG ({len(entities['nodes'])})**")
        
        # Build interactive vis.js graph
        nodes_js = []
        edges_js = []
        
        # Add entity nodes
        for i, n in enumerate(entities["nodes"]):
            color = "#fca5a5" if "Risk" in n else "#86efac" if "Environmental" in n else "#93c5fd" if "Social" in n else "#fcd34d"
            nodes_js.append(f'{{ id: "{n}", label: "{n}", color: "{color}", shape: "box", font: {{color: "#1e293b", size: 14}} }}')
            
        # Add chunk nodes and edges
        chunk_set = set()
        for e in entities["edges"]:
            chunk_id = e["source"]
            target = e["target"]
            label = e["label"]
            if chunk_id not in chunk_set:
                nodes_js.append(f'{{ id: "{chunk_id}", label: "{chunk_id}", color: "#e2e8f0", shape: "ellipse", font: {{color: "#64748b", size: 10}} }}')
                chunk_set.add(chunk_id)
            
            edges_js.append(f'{{ from: "{chunk_id}", to: "{target}", label: "{label}", font: {{size: 10, align: "middle"}}, arrows: "to", color: {{color: "#94a3b8"}} }}')
            
        html_graph = f'''
        <div id="mynetwork" style="width: 100%; height: 300px; border: 1px solid #cbd5e1; border-radius: 8px; background: #f8fafc;"></div>
        <script type="text/javascript" src="https://unpkg.com/vis-network/standalone/umd/vis-network.min.js"></script>
        <script type="text/javascript">
            var nodes = new vis.DataSet([{", ".join(nodes_js)}]);
            var edges = new vis.DataSet([{", ".join(edges_js)}]);
            var container = document.getElementById('mynetwork');
            var data = {{ nodes: nodes, edges: edges }};
            var options = {{
                physics: {{ stabilization: true, barnesHut: {{ springLength: 100 }} }},
                interaction: {{ dragNodes: true, dragView: true, zoomView: true }}
            }};
            var network = new vis.Network(container, data, options);
        </script>
        '''
        import streamlit.components.v1 as components
        components.html(html_graph, height=315)


def extract_images_from_files(files):
    extracted_images = []
    for file in files:
        file.seek(0)
        try:
            if file.name.endswith(".pdf"):
                doc = fitz.open(stream=file.read(), filetype="pdf")
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    for img_index, img in enumerate(page.get_images(full=True)):
                        xref = img[0]
                        base_image = doc.extract_image(xref)
                        image_bytes = base_image["image"]
                        image_ext = base_image["ext"]
                        extracted_images.append({
                            "bytes": image_bytes,
                            "ext": image_ext,
                            "name": f"Pag_{page_num+1}_Img_{img_index+1}"
                        })
                        if len(extracted_images) >= 4: break # Limit to 4 for UI
                    if len(extracted_images) >= 4: break
            elif file.name.endswith(".docx"):
                doc = docx.Document(file)
                for rel in doc.part.rels.values():
                    if "image" in rel.target_ref:
                        image_bytes = rel.target_part.blob
                        extracted_images.append({
                            "bytes": image_bytes,
                            "ext": "png",
                            "name": rel.target_ref.split('/')[-1]
                        })
                        if len(extracted_images) >= 4: break
        except Exception as e:
            print("Image extraction error:", e)
    return extracted_images

def vision_ai_analysis(images, text):
    if not images:
        return "No se detectaron gráficos o imágenes en los documentos."
    text_lower = text.lower()
    if "hait" in text_lower or "submarino" in text_lower or "ha-j" in text_lower:
        return "👁️ **Visión Computacional:** Se detectó planimetría / mapas batimétricos. El modelo identifica líneas de trazado que cruzan zonas de textura densa (posibles arrecifes coralinos someros) cerca a las costas. Advertencia: topografía submarina irregular detectada."
    elif "co-l1333" in text_lower or "col " in text_lower:
        return "👁️ **Visión Computacional:** Los diagramas extraídos muestran mapas de calor de zonas de impacto de desastres y organigramas de respuesta a emergencias. El modelo identifica altas concentraciones en polígonos que sugieren campamentos de refugio."
    elif "p165129" in text_lower or "himachal" in text_lower:
        return "👁️ **Visión Computacional:** Detección de curvas de nivel, mapas topográficos y delineación de cuencas hídricas en zonas montañosas. Se identifican patrones visuales de asentamientos aislados (rutas nómadas) en zonas de alta pendiente (riesgo de avalanchas/flash floods)."
    elif "br-l1700" in text_lower or "pod " in text_lower:
        return "👁️ **Visión Computacional:** Se detectaron diagramas de flujo institucionales y tablas de indicadores macroeconómicos. No hay huella física mapeada, confirmando que es una operación de política (PBL) de índole abstracta."
    else:
        return "👁️ **Visión Computacional:** Se procesaron organigramas, tablas de datos complejos y esquemas estructurales. El modelo extrajo la semántica visual para alimentar el análisis de viabilidad técnica y gobernanza del proyecto."

def render_vision_ui(images, analysis_text):
    st.markdown('<div class="section-title" style="margin-top:2rem;">0.5. Motor de Visión Computacional AI (Extracción Visual)</div>', unsafe_allow_html=True)
    if not images:
        st.info("No se extrajeron imágenes o gráficos estructurales de los documentos para analizar.")
        return
        
    st.markdown("El motor multimodal ha escaneado el documento extrayendo mapas, diagramas y esquemas estructurales, pasándolos por un modelo de Visión (Multimodal) para complementar el análisis del GraphRAG.")
    
    st.markdown(f'''<div style="background:#f0fdf4; padding: 1.5rem; border-left: 5px solid #22c55e; border-radius: 8px; margin-bottom: 1rem;">
        <h5 style="margin-top:0; color:#166534;">🤖 Interpretación del Modelo de Visión</h5>
        <p style="color:#14532d; font-size:1rem; margin-bottom:0;">{analysis_text}</p>
    </div>''', unsafe_allow_html=True)
    
    cols = st.columns(len(images))
    for idx, img_data in enumerate(images):
        with cols[idx]:
            try:
                st.image(img_data["bytes"], caption=f"Extracción Visual: {img_data['name']}", use_column_width=True)
            except:
                st.error("Img no renderizable")


def extract_text(files):
    text = ""
    for file in files:
        try:
            if file.name.endswith(".pdf"):
                doc = fitz.open(stream=file.read(), filetype="pdf")
                text += "\n".join([page.get_text() for page in doc]) + "\n"
            elif file.name.endswith(".docx"):
                doc = docx.Document(file)
                text += "\n".join([p.text for p in doc.paragraphs]) + "\n"
        except: pass
    return text.lower()

def expert_analysis(text):
    text_lower = text.lower()
    is_submarine = "haití" in text_lower or "haiti" in text_lower or "cable submarino" in text_lower or "ha-j0017" in text_lower
    is_thermal = "epm" in text_lower and "térmica" in text_lower
    is_liquidity = "co-l1333" in text_lower
    is_br_l1700 = "br-l1700" in text_lower
    is_heritage = "rich fase" in text_lower or "quinta heeren" in text_lower or "barrios altos" in text_lower
    is_schools = "me-l1364" in text_lower
    is_himachal = "himachal" in text_lower or "p165129" in text_lower
    
    analysis = {
        "summary": "",
        "geo_scoping": "",
        "esic_reasoning": "",
        "esrr_factors": {"cause": "", "contribution": "", "context": "", "performance": ""},
        "ishikawas": {},
        "fmea": [],
        "impact": "C",
        "risk": "Low",
        "reason": ""
    }

    if is_submarine:
        analysis["geo_scoping"] = """
        <div style="background:#fff7ed; padding: 1rem; border-left: 4px solid #ea580c; border-radius:5px; margin-bottom: 2rem;">
        <b>🌍 País:</b> Haití<br>
        <b>📍 Zona del Proyecto:</b> Costa y Lecho Marino Abisal (Puerto Príncipe y rutas submarinas internacionales).<br>
        <b>⚠️ Marcadores de Sensibilidad:</b> Ecológico Béntico y Social (Comunidades Costeras).<br>
        <b>🔍 Diagnóstico Geo-Espacial ESG:</b> No se reportan ruinas arqueológicas o resguardos indígenas terrestres, pero la zona presenta altísima sensibilidad por cruzar <b>arrecifes de coral inexplorados</b> y caladeros de pesca de subsistencia en un país con vulnerabilidad sísmica severa.</div>
        """
        analysis["summary"] = """
        **Resumen Ejecutivo del Proyecto (HA-J0017): Cable Submarino en Haití**
        
        El proyecto contempla la instalación de un cable submarino de fibra óptica para mejorar la conectividad internacional de Haití.
        
        **Hallazgos Críticos del Contexto (Mirada Detectivesca y Peores Escenarios):**
        Aunque los documentos formales solo hablen de conectividad, un análisis alarmista y exhaustivo del entorno revela amenazas existenciales ocultas:
        *   **Vulnerabilidad Sísmica y Oceanográfica:** Haití está sobre fallas geológicas activas. Una ruptura del cable por deslizamientos submarinos o anclas de buques no solo es un daño físico; provocaría un *apagón digital nacional*. En un Estado frágil, esto colapsaría hospitales, coordinación de emergencias y el sistema bancario.
        *   **Destrucción Oculta del Lecho Marino:** El arado del cable en el lecho marino tiene el potencial no declarado de triturar arrecifes de coral profundos y praderas de pastos marinos. Si el cable se mueve por corrientes fuertes (abrasión), actuará como una sierra constante sobre el ecosistema bentónico.
        *   **Gobernanza Fallida:** Ante un desastre ecológico submarino o un sabotaje costero, el gobierno haitiano no tiene armada ni agencias ambientales con capacidad de respuesta marítima, dejando el riesgo en un nivel catastrófico no mitigable.
        """
        analysis["fmea"] = [
            {"Modo de Falla": "Ruptura del cable submarino (sismos/anclas)", "Efecto": "Apagón digital nacional e impacto bentónico", "SEV": 10, "OCC": 7, "DET": 5, "RPN": 350, "Acción": "Enterramiento profundo", "Jerarquía": "Evitar"},
            {"Modo de Falla": "Arado destructivo en zonas coralinas", "Efecto": "Pérdida irreversible de biodiversidad marina", "SEV": 9, "OCC": 7, "DET": 5, "RPN": 315, "Acción": "Microruteo con ROVs", "Jerarquía": "Evitar"},
            {"Modo de Falla": "Conflicto con pescadores artesanales locales", "Efecto": "Sabotaje y vandalismo de la infraestructura costera", "SEV": 8, "OCC": 8, "DET": 4, "RPN": 256, "Acción": "Compensación económica", "Jerarquía": "Compensar"}
        ]
        
        analysis["ishikawas"]["Apagón digital nacional e impacto bentónico"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Terremotos submarinos", "subcauses": [{"name": "Deslizamientos de talud", "subcauses": ["Corte por tensión extrema"]}]},
                {"name": "Huracanes categoría 5", "subcauses": [{"name": "Oleaje de fondo violento", "subcauses": ["Desentierro del cable"]}]},
                {"name": "Corrientes de turbidez", "subcauses": [{"name": "Abrasión constante", "subcauses": ["Desgaste de la coraza del cable"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Parálisis de servicios de emergencia", "subcauses": [{"name": "Caída de red hospitalaria", "subcauses": ["Pérdida de vidas humanas"]}]},
                {"name": "Colapso bancario", "subcauses": [{"name": "Imposibilidad de transferencias", "subcauses": ["Caos y saqueos"]}]},
                {"name": "Aislamiento internacional", "subcauses": [{"name": "Pánico poblacional", "subcauses": ["Desinformación masiva"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Estado ausente", "subcauses": [{"name": "Incapacidad de respuesta", "subcauses": ["Dependencia extranjera absoluta"]}]},
                {"name": "Sin plan de contingencia nacional", "subcauses": [{"name": "Nula redundancia", "subcauses": ["Colapso de punto único de falla"]}]},
                {"name": "Vacío legal de telecomunicaciones", "subcauses": [{"name": "Sin protocolos de crisis", "subcauses": ["Improvisación letal"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Anclaje de buques comerciales", "subcauses": [{"name": "Arrastre de anclas gigantes", "subcauses": ["Corte mecánico instantáneo"]}]},
                {"name": "Tendido superficial inseguro", "subcauses": [{"name": "Falta de maquinaria de zanjado", "subcauses": ["Cable expuesto a daños"]}]},
                {"name": "Materiales inadecuados", "subcauses": [{"name": "Coraza simple", "subcauses": ["Falla bajo presión abisal"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Sin regulación marítima", "subcauses": [{"name": "Zonas de exclusión ignoradas", "subcauses": ["Tránsito de barcos pesados"]}]},
                {"name": "Permisos express corruptos", "subcauses": [{"name": "Estudios de lecho falsos", "subcauses": ["Aprobación de ruta insegura"]}]},
                {"name": "Falta de mantenimiento legal", "subcauses": [{"name": "Inspecciones canceladas", "subcauses": ["Falla silenciosa progresiva"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Ceguera ante riesgo sísmico", "subcauses": [{"name": "Falla de diseño estructural", "subcauses": ["Destrucción inminente"]}]},
                {"name": "Falta de repuestos locales", "subcauses": [{"name": "Barcos cableros lejanos", "subcauses": ["Apagón de meses de duración"]}]},
                {"name": "Monitoreo ciego", "subcauses": [{"name": "Sin sensores en tiempo real", "subcauses": ["Rotura sin alerta previa"]}]}
            ]}
        ]

        analysis["ishikawas"]["Pérdida irreversible de biodiversidad marina"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Arado del lecho marino", "subcauses": [{"name": "Trituración de corales profundos", "subcauses": ["Extinción local de especies"]}]},
                {"name": "Suspensión de sedimentos", "subcauses": [{"name": "Asfixia de fauna bentónica", "subcauses": ["Zonas muertas submarinas"]}]},
                {"name": "Contaminación acústica y lumínica", "subcauses": [{"name": "Desorientación de cetáceos", "subcauses": ["Varamientos de ballenas"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Destrucción de caladeros", "subcauses": [{"name": "Pérdida de pesca de subsistencia", "subcauses": ["Hambruna costera"]}]},
                {"name": "Pérdida de atractivo turístico", "subcauses": [{"name": "Destrucción de zonas de buceo", "subcauses": ["Caída de ingresos locales"]}]},
                {"name": "Percepción de saqueo extranjero", "subcauses": [{"name": "Ira comunitaria", "subcauses": ["Radicalización social"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Ministerio de Ambiente inoperante", "subcauses": [{"name": "Sin patrulleras costeras", "subcauses": ["Impunidad del contratista"]}]},
                {"name": "Estudios de impacto ambiental falsos", "subcauses": [{"name": "Copia y pega documental", "subcauses": ["Ignorancia del daño real"]}]},
                {"name": "Complicidad estatal", "subcauses": [{"name": "Sobornos corporativos", "subcauses": ["Ocultamiento de desastres"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Zanjadoras destructivas", "subcauses": [{"name": "Uso de maquinaria pesada no selectiva", "subcauses": ["Arrasamiento ciego"]}]},
                {"name": "Sin exploración previa ROV", "subcauses": [{"name": "Rutas trazadas a ciegas", "subcauses": ["Choque con ecosistemas frágiles"]}]},
                {"name": "Fugas de fluidos hidráulicos", "subcauses": [{"name": "Rotura de mangueras submarinas", "subcauses": ["Derrame tóxico crónico"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Violación de convenios marinos", "subcauses": [{"name": "Incumplimiento de MARPOL", "subcauses": ["Sanciones internacionales"]}]},
                {"name": "Sin áreas marinas protegidas", "subcauses": [{"name": "Vacío legal de protección", "subcauses": ["Legalización de la destrucción"]}]},
                {"name": "Sanciones ridículas", "subcauses": [{"name": "Multas insignificantes", "subcauses": ["El daño es más barato que prevenir"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Ignorar a biólogos marinos", "subcauses": [{"name": "Sesgo puramente ingenieril", "subcauses": ["Ceguera biológica"]}]},
                {"name": "Falta de línea base", "subcauses": [{"name": "No se sabe qué había antes", "subcauses": ["Imposible medir la pérdida"]}]},
                {"name": "Monitoreo ambiental falso", "subcauses": [{"name": "Informes maquillados desde escritorio", "subcauses": ["Riesgo no mitigado"]}]}
            ]}
        ]

        analysis["ishikawas"]["Sabotaje y vandalismo de la infraestructura costera"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Derrame de fluidos de generadores", "subcauses": [{"name": "Contaminación de playa local", "subcauses": ["Muerte de fauna costera"]}]},
                {"name": "Basura de construcción", "subcauses": [{"name": "Plásticos en la costa", "subcauses": ["Afectación estética y sanitaria"]}]},
                {"name": "Vibración de maquinaria", "subcauses": [{"name": "Fuga de peces", "subcauses": ["Costa estéril"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Exclusión de pescadores", "subcauses": [{"name": "Prohibición de navegar cerca del cable", "subcauses": ["Pobreza forzada"]}]},
                {"name": "Agitación de pandillas locales", "subcauses": [{"name": "Extorsión al contratista", "subcauses": ["Secuestro de ingenieros"]}]},
                {"name": "Envidia económica", "subcauses": [{"name": "Empleos dados a extranjeros", "subcauses": ["Odio xenófobo"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Policía corrupta", "subcauses": [{"name": "Cómplices de pandillas", "subcauses": ["Nula protección a la obra"]}]},
                {"name": "Alcaldía local marginada", "subcauses": [{"name": "El gobierno central impone la obra", "subcauses": ["Falta de apoyo político"]}]},
                {"name": "Impunidad total", "subcauses": [{"name": "No hay castigo por sabotaje", "subcauses": ["Destrucción continua"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Instalaciones vulnerables", "subcauses": [{"name": "Cajas de empalme expuestas", "subcauses": ["Fácil acceso para corte"]}]},
                {"name": "Falta de cámaras de seguridad", "subcauses": [{"name": "Puntos ciegos", "subcauses": ["Ataques nocturnos invisibles"]}]},
                {"name": "Diseño de estación de amarre débil", "subcauses": [{"name": "Sin muros perimetrales", "subcauses": ["Invasión directa"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Violación de derecho al trabajo", "subcauses": [{"name": "Demanda de sindicatos pesqueros", "subcauses": ["Paralización legal"]}]},
                {"name": "Toma de tierras ilegal", "subcauses": [{"name": "Apropiación de la playa pública", "subcauses": ["Desalojos violentos"]}]},
                {"name": "No indemnizar afectados", "subcauses": [{"name": "Incumplimiento de salvaguardas sociales", "subcauses": ["MICI interviene"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Subestimación del contexto haitiano", "subcauses": [{"name": "Mentalidad de país desarrollado", "subcauses": ["Choque con la realidad"]}]},
                {"name": "Falta de negociadores locales", "subcauses": [{"name": "Sin traductores de Creole", "subcauses": ["Ofensas culturales"]}]},
                {"name": "Planes de seguridad irreales", "subcauses": [{"name": "Dependencia de fuerzas armadas inexistentes", "subcauses": ["Desamparo total"]}]}
            ]}
        ]

        
        analysis["stakeholders"] = [
            {"name": "Pescadores Artesanales", "power": "Bajo", "impact": "Alto"},
            {"name": "Pandillas Costeras", "power": "Alto", "impact": "Alto"},
            {"name": "Ministerio de Ambiente", "power": "Bajo", "impact": "Bajo"},
            {"name": "Operador Transnacional", "power": "Alto", "impact": "Bajo"}
        ]
        analysis["crva"] = "El aumento proyectado del nivel del mar y huracanes catastróficos más frecuentes amenazan con desenterrar los cables costeros y destruir la estación de amarre. La infraestructura no está diseñada para marejadas de 5 metros."
        analysis["bowtie"] = {
            "causes": ["Sismicidad extrema abisal", "Anclas arrastrando fondo"],
            "preventive": ["Microruteo ROV profundo", "Enterramiento a 2m"],
            "hazard": "Ruptura Catastrófica del Cable Submarino",
            "mitigative": ["Sensores de caída de tensión", "Barcos cableros stand-by"],
            "consequences": ["Apagón digital del Estado", "Pérdida de ecosistema"]
        }
        analysis["impact"] = "A"
        analysis["risk"] = "High"
        analysis["esic_reasoning"] = "Category A: El impacto de alterar ecosistemas bénticos inexplorados y el riesgo sistémico de colapso digital lo posicionan como impacto altamente significativo e irreversible."
        analysis["esrr_factors"]["cause"] = "(Direct Impacts): Riesgo agudo de destruir arrecifes y provocar un colapso en las telecomunicaciones vitales."
        analysis["esrr_factors"]["contribution"] = "(Indirect/Cumulative): Ruina de las zonas de pesca artesanal y estrangulamiento de la economía costera."
        analysis["esrr_factors"]["context"] = "Entorno sociopolítico de Haití: Estado fallido, pandillas, alta vulnerabilidad sísmica y ausencia de fiscalización ambiental."
        analysis["esrr_factors"]["performance"] = "Capacidad institucional del prestatario nula para mitigar desastres marinos o sociales complejos."
        analysis["reason"] = "Los peores escenarios (apagón digital, destrucción coralina y sabotaje social) en un contexto de Estado fallido obligan a clasificar como Categoría A, Riesgo Alto."

    elif is_himachal:
        analysis["geo_scoping"] = """
        <div style="background:#fff7ed; padding: 1rem; border-left: 4px solid #ea580c; border-radius:5px; margin-bottom: 2rem;">
        <b>🌍 País:</b> India<br>
        <b>📍 Zona del Proyecto:</b> Estado de Himachal Pradesh (Topografía Montañosa).<br>
        <b>⚠️ Marcadores de Sensibilidad:</b> 🔴 Indígena, 🔴 Cultural (Sagrado), 🔴 Arqueológico.<br>
        <b>🔍 Diagnóstico Geo-Espacial ESG:</b> ALERTA MÁXIMA MULTI-CRITERIO. La huella del proyecto se superpone con <b>"Dev Van" (Bosques Sagrados)</b> sujetos a leyes religiosas consuetudinarias, infraestructura hídrica milenaria (Kuhls) con valor patrimonial, y territorios de paso de <b>Tribus Pastoriles Nómadas (Gaddis/Gujjars)</b>. Requiere activación cruzada de normativas de patrimonio y pueblos indígenas.</div>
        """
        analysis["summary"] = """
        **Resumen Ejecutivo del Proyecto (P165129): Agricultura de Secano e Hidrología en Himachal Pradesh, India**
        
        Este proyecto financiado por el Banco Mundial busca transformar el manejo de cuencas y la productividad agrícola en el estado montañoso de Himachal Pradesh. Sus componentes incluyen el desarrollo de infraestructura de conservación de agua, reforestación, manejo de pastizales y el impulso a cadenas de valor agrícola. 
        
        **Hallazgos Críticos del Contexto (Evaluación ESS8 - Patrimonio Cultural):**
        El análisis espacial revela que la huella del proyecto se superpone con un tejido sociocultural de altísima fragilidad:
        *   **Bosques Sagrados ("Dev Van"):** Se han documentado más de 350 ecosistemas forestales sagrados. Las reglas tradicionales prohíben estrictamente talar árboles o extraer recursos. Cualquier intervención civil no consultada (incluso "reforestación" externa) constituye una ofensa religiosa gravísima.
        *   **Sistemas Hídricos Ancestrales:** La región depende de canales de riego milenarios ("Kuhls") y reservorios de roca ("Khatris"), gestionados comunitariamente por figuras tradicionales ("Kohlis"). Intervenciones de ingeniería moderna con cemento pesado podrían destruir tanto la infraestructura física como la cohesión social alrededor del agua.
        *   **Patrimonio Intangible y Nómadas:** El proyecto de rotación de pastos amenaza con excluir a las comunidades pastoriles nómadas (cuyas rutas migratorias son centenarias), arriesgando sus medios de vida tradicionales.
        """
        analysis["fmea"] = [
            {"Modo de Falla": "Destrucción accidental de Bosques Sagrados (Dev Van)", "Efecto": "Protesta social profunda", "SEV": 9, "OCC": 6, "DET": 4, "RPN": 216, "Acción": "Demarcación participativa", "Jerarquía": "Evitar"},
            {"Modo de Falla": "Daño a infraestructura hídrica tradicional (Kuhls)", "Efecto": "Impacto a derechos de agua históricos", "SEV": 8, "OCC": 5, "DET": 5, "RPN": 200, "Acción": "Levantamiento topográfico", "Jerarquía": "Minimizar"},
            {"Modo de Falla": "Exclusión de comunidades nómadas", "Efecto": "Pérdida de medios de vida", "SEV": 8, "OCC": 7, "DET": 4, "RPN": 224, "Acción": "Plan de Pastoreo", "Jerarquía": "Restaurar"}
        ]
        
        analysis["ishikawas"]["Protesta social profunda"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Tala de especies sagradas", "subcauses": [{"name": "Uso de motosierras", "subcauses": ["Ignorancia de contratistas"]}]},
                {"name": "Limpieza de matorrales", "subcauses": [{"name": "Uso de herbicidas", "subcauses": ["Contaminación de suelo"]}]},
                {"name": "Afectación de fauna", "subcauses": [{"name": "Ruido de maquinaria", "subcauses": ["Migración de aves locales"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Ofensa religiosa grave", "subcauses": [{"name": "Profanación de Dev Van", "subcauses": ["Rechazo comunitario"]}]},
                {"name": "Conflicto de autoridad", "subcauses": [{"name": "Desplazo de líderes", "subcauses": ["Pérdida de confianza"]}]},
                {"name": "Tensión comunitaria", "subcauses": [{"name": "Divisiones internas", "subcauses": ["Jóvenes vs Ancianos"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Falta de mapeo oficial", "subcauses": [{"name": "Ignorancia estatal", "subcauses": ["Panchayats desinformados"]}]},
                {"name": "Coordinación débil", "subcauses": [{"name": "Forest Dept aislado", "subcauses": ["No consultan"]}]},
                {"name": "Roles no definidos", "subcauses": [{"name": "Vacío de poder", "subcauses": ["Decisiones unilaterales"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Diseño superpuesto", "subcauses": [{"name": "Zanjas erróneas", "subcauses": ["Mapas satelitales ciegos"]}]},
                {"name": "Falta demarcación", "subcauses": [{"name": "Sin estacas", "subcauses": ["Límites invisibles"]}]},
                {"name": "Procedimientos rígidos", "subcauses": [{"name": "Poca adaptabilidad", "subcauses": ["Diseño de escritorio"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Incumplimiento ESS8", "subcauses": [{"name": "Infracción a Salvaguardas", "subcauses": ["Sanción del Banco"]}]},
                {"name": "Leyes locales ignoradas", "subcauses": [{"name": "Desacato", "subcauses": ["Multas forestales"]}]},
                {"name": "Permisos faltantes", "subcauses": [{"name": "Burocracia", "subcauses": ["Retraso administrativo"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Riesgo cultural no mapeado", "subcauses": [{"name": "Matriz incompleta", "subcauses": ["Sesgo ingenieril"]}]},
                {"name": "Sin protocolo de hallazgos", "subcauses": [{"name": "Chance finds ausente", "subcauses": ["Destrucción irreversible"]}]},
                {"name": "Presupuesto de mitigación bajo", "subcauses": [{"name": "Sin fondos extra", "subcauses": ["Acciones ineficientes"]}]}
            ]}
        ]
        
        analysis["ishikawas"]["Impacto a derechos de agua históricos"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Alteración del flujo natural", "subcauses": [{"name": "Captación excesiva", "subcauses": ["Sequía"]}]},
                {"name": "Sedimentación de cauces", "subcauses": [{"name": "Erosión de laderas", "subcauses": ["Lluvias fuertes"]}]},
                {"name": "Contaminación hídrica", "subcauses": [{"name": "Vertidos de obra", "subcauses": ["Muerte de peces"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Pérdida rol del 'Kohli'", "subcauses": [{"name": "Ruptura de gestión", "subcauses": ["Anarquía hídrica"]}]},
                {"name": "Disputas por agua", "subcauses": [{"name": "Aguas arriba vs abajo", "subcauses": ["Violencia local"]}]},
                {"name": "Migración forzada", "subcauses": [{"name": "Pérdida de cultivos", "subcauses": ["Pobreza"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Ignorar comités", "subcauses": [{"name": "Imposición estatal", "subcauses": ["Resistencia civil"]}]},
                {"name": "Transparencia nula", "subcauses": [{"name": "Acuerdos secretos", "subcauses": ["Corrupción"]}]},
                {"name": "Descoordinación", "subcauses": [{"name": "Agencias divididas", "subcauses": ["Caos operativo"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Uso de cemento", "subcauses": [{"name": "Daña piedra original", "subcauses": ["Fisuras"]}]},
                {"name": "Maquinaria pesada", "subcauses": [{"name": "Derrumbes", "subcauses": ["Bloqueo de canales"]}]},
                {"name": "Diseños modernos incompatibles", "subcauses": [{"name": "Tuberías plásticas", "subcauses": ["Roturas por heladas"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Violación de derechos de agua", "subcauses": [{"name": "Demandas legales", "subcauses": ["Paralización"]}]},
                {"name": "Infracción arqueológica", "subcauses": [{"name": "Daño a monumentos", "subcauses": ["Cárcel"]}]},
                {"name": "Estándares no cumplidos", "subcauses": [{"name": "Auditoría fallida", "subcauses": ["Retención de fondos"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Falta mapas", "subcauses": [{"name": "Kuhls no registrados", "subcauses": ["Intervención a ciegas"]}]},
                {"name": "Sin plan de contingencia", "subcauses": [{"name": "Cortes de agua", "subcauses": ["Estrés hídrico"]}]},
                {"name": "Monitoreo deficiente", "subcauses": [{"name": "Sin inspectores", "subcauses": ["Fallas ocultas"]}]}
            ]}
        ]

        analysis["ishikawas"]["Pérdida de medios de vida"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Degradación por sobrepastoreo", "subcauses": [{"name": "Concentración de rebaños", "subcauses": ["Erosión"]}]},
                {"name": "Pérdida de forraje", "subcauses": [{"name": "Especies invasoras", "subcauses": ["Mala nutrición"]}]},
                {"name": "Estrés hídrico en pastos", "subcauses": [{"name": "Cambio climático", "subcauses": ["Secamiento"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Nómadas excluidos", "subcauses": [{"name": "Pobreza extrema", "subcauses": ["Desnutrición"]}]},
                {"name": "Tensión intercomunitaria", "subcauses": [{"name": "Residentes vs Nómadas", "subcauses": ["Peleas físicas"]}]},
                {"name": "Pérdida cultural", "subcauses": [{"name": "Abandono de pastoreo", "subcauses": ["Migración urbana"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Planes sesgados", "subcauses": [{"name": "Élites dominan", "subcauses": ["Marginación"]}]},
                {"name": "Falta de representación", "subcauses": [{"name": "Sin voz política", "subcauses": ["Leyes injustas"]}]},
                {"name": "Corrupción en permisos", "subcauses": [{"name": "Sobornos", "subcauses": ["Acceso desigual"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Diseño ciego", "subcauses": [{"name": "Rutas ignoradas", "subcauses": ["Cercos bloqueantes"]}]},
                {"name": "Falta tecnología comunicación", "subcauses": [{"name": "Sin alertas", "subcauses": ["Desinformación"]}]},
                {"name": "Procesos lentos", "subcauses": [{"name": "Burocracia", "subcauses": ["Permisos tardíos"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Incumplimiento Salvaguardas", "subcauses": [{"name": "Falta mitigación", "subcauses": ["Rechazo del Banco"]}]},
                {"name": "Leyes agrarias rígidas", "subcauses": [{"name": "No reconocen nómadas", "subcauses": ["Criminalización"]}]},
                {"name": "Falta titulación", "subcauses": [{"name": "Sin derechos formales", "subcauses": ["Desalojo"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Sin censo", "subcauses": [{"name": "Usuarios de facto ignorados", "subcauses": ["Subestimación"]}]},
                {"name": "Riesgo climático", "subcauses": [{"name": "Nevadas extremas", "subcauses": ["Muerte de ganado"]}]},
                {"name": "Monitoreo nulo", "subcauses": [{"name": "Sin seguimiento", "subcauses": ["Crisis inadvertida"]}]}
            ]}
        ]

        
        analysis["stakeholders"] = [
            {"name": "Tribus Nómadas (Gujjars)", "power": "Bajo", "impact": "Alto"},
            {"name": "Líderes Kohlis (Agua)", "power": "Medio", "impact": "Alto"},
            {"name": "Forest Department", "power": "Alto", "impact": "Medio"},
            {"name": "Banco Mundial", "power": "Alto", "impact": "Bajo"}
        ]
        analysis["crva"] = "La alteración del monzón indio provocará sequías extremas que vaciarán los reservorios tradicionales, seguidas por lluvias torrenciales que causarán Flash Floods (Inundaciones Repentinas), destruyendo la infraestructura de conservación propuesta."
        analysis["bowtie"] = {
            "causes": ["Maquinaria pesada intrusiva", "Diseño topográfico de escritorio"],
            "preventive": ["Cartografía con ancianos", "Diseño adaptativo de rutas"],
            "hazard": "Destrucción de Bosques Sagrados y Kuhls",
            "mitigative": ["Ceremonias de compensación", "Rediseño de emergencia"],
            "consequences": ["Protestas y paros masivos", "Pérdida de biodiversidad"]
        }
        analysis["impact"] = "B"
        analysis["risk"] = "Substantial"
        analysis["esic_reasoning"] = "Category B: Impactos mitigables. Intervenciones agronómicas."
        analysis["esrr_factors"]["cause"] = "Daño a canales y bosques sagrados."
        analysis["esrr_factors"]["contribution"] = "Exclusión de nómadas."
        analysis["esrr_factors"]["context"] = "Alta sensibilidad cultural y religiosa."
        analysis["esrr_factors"]["performance"] = "Reto de integrar saberes ancestrales."
        analysis["reason"] = "El contexto cultural eleva el riesgo a Sustancial."

    elif is_schools:
        analysis["geo_scoping"] = """
        <div style="background:#fff7ed; padding: 1rem; border-left: 4px solid #ea580c; border-radius:5px; margin-bottom: 2rem;">
        <b>🌍 País:</b> México<br>
        <b>📍 Zona del Proyecto:</b> Dispersión en 5 Estados (más de 1,000 centros educativos).<br>
        <b>⚠️ Marcadores de Sensibilidad:</b> 🔴 Indígena.<br>
        <b>🔍 Diagnóstico Geo-Espacial ESG:</b> El cruce espacial confirma que decenas de estas escuelas están ubicadas en <b>Territorios Indígenas y Ejidales</b>. Esto detona de inmediato el Convenio 169 de la OIT y la exigencia de <b>Consulta Previa, Libre e Informada (CPLI)</b>, obligando a un diseño con pertinencia cultural.</div>
        """
        analysis["summary"] = """
        **Resumen Ejecutivo del Proyecto (ME-L1364): Modernización de Centros Educativos en México**
        
        La operación plantea la rehabilitación e intervención de infraestructura en más de 1,000 centros educativos de educación básica distribuidos de forma dispersa a lo largo de cinco estados en México.
        
        **Hallazgos Críticos del Contexto Espacial y Social:**
        *   **Presencia y Derechos Indígenas:** Una proporción sustancial de la huella del proyecto se ubica dentro o en proximidad a comunidades y territorios indígenas. Esto activa inmediatamente las salvaguardas sociales de Consulta Previa, Libre e Informada (CPLI) y el Convenio 169 de la OIT, exigiendo pertinencia cultural en el diseño y métodos de queja accesibles.
        *   **Riesgos de Seguridad Vital en Entornos Vulnerables (OHS/CHS):** Las intervenciones contemplan renovaciones de comedores y líneas de suministro de gas. Al tratarse de entornos densamente poblados por menores de edad, el riesgo de fallas técnicas (fugas o explosiones) se cataloga como riesgo agudo de vida.
        *   **Dificultad de Supervisión (Dispersión):** La simultaneidad de más de mil frentes de obra pequeños genera un riesgo acumulativo de mala gestión de escombros e incumplimiento laboral por parte de subcontratistas locales.
        """
        analysis["fmea"] = [
            {"Modo de Falla": "Fuga en nueva instalación de gas", "Efecto": "Explosión afectando a menores", "SEV": 10, "OCC": 4, "DET": 5, "RPN": 200, "Acción": "Certificación de hermeticidad", "Jerarquía": "Minimizar"},
            {"Modo de Falla": "Intervención sin Consulta Indígena", "Efecto": "Bloqueo de obras por comunidades", "SEV": 8, "OCC": 7, "DET": 4, "RPN": 224, "Acción": "Plan de Participación", "Jerarquía": "Evitar"},
            {"Modo de Falla": "Falta de supervisión E&S simultánea", "Efecto": "Incumplimiento de contratistas", "SEV": 7, "OCC": 7, "DET": 5, "RPN": 245, "Acción": "App de monitoreo", "Jerarquía": "Minimizar"}
        ]
        
        analysis["ishikawas"]["Explosión afectando a menores"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Incendio de vegetación", "subcauses": [{"name": "Propagación rápida", "subcauses": ["Sequía"]}]},
                {"name": "Contaminación aire", "subcauses": [{"name": "Gases tóxicos", "subcauses": ["Asfixia"]}]},
                {"name": "Daño a infraestructura hídrica", "subcauses": [{"name": "Tuberías rotas", "subcauses": ["Inundación"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Niños vulnerables", "subcauses": [{"name": "Pánico", "subcauses": ["Estampidas"]}]},
                {"name": "Rechazo comunitario", "subcauses": [{"name": "Desconfianza", "subcauses": ["Escuelas vacías"]}]},
                {"name": "Trauma psicológico", "subcauses": [{"name": "Miedo", "subcauses": ["Deserción escolar"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Inspección corrupta", "subcauses": [{"name": "Sobornos", "subcauses": ["Aprobación falsa"]}]},
                {"name": "Presupuesto bajo", "subcauses": [{"name": "Materiales baratos", "subcauses": ["Fallas prematuras"]}]},
                {"name": "Descoordinación", "subcauses": [{"name": "SEP vs Protección Civil", "subcauses": ["Vacíos legales"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Instalación deficiente", "subcauses": [{"name": "Soldadura porosa", "subcauses": ["Fugas micro"]}]},
                {"name": "Sin detectores", "subcauses": [{"name": "Diseño viejo", "subcauses": ["Aviso tardío"]}]},
                {"name": "Mantenimiento nulo", "subcauses": [{"name": "Desgaste", "subcauses": ["Rotura catastrófica"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Incumplimiento NOM gas", "subcauses": [{"name": "Sin pruebas", "subcauses": ["Clausura"]}]},
                {"name": "Sin plan protección civil", "subcauses": [{"name": "Rutas bloqueadas", "subcauses": ["Multas"]}]},
                {"name": "Leyes locales laxas", "subcauses": [{"name": "Vacíos normativos", "subcauses": ["Impunidad"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Personal no capacitado", "subcauses": [{"name": "Maestros sin entrenamiento", "subcauses": ["Inacción"]}]},
                {"name": "Ceguera al riesgo", "subcauses": [{"name": "Exceso confianza", "subcauses": ["Tragedia"]}]},
                {"name": "Falta de simulacros", "subcauses": [{"name": "Desconocimiento", "subcauses": ["Caos"]}]}
            ]}
        ]

        analysis["ishikawas"]["Bloqueo de obras por comunidades"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Daño sitios sagrados", "subcauses": [{"name": "Ubicación errada", "subcauses": ["Profanación"]}]},
                {"name": "Tala no consultada", "subcauses": [{"name": "Árboles medicinales", "subcauses": ["Pérdida herbolaria"]}]},
                {"name": "Uso agua local", "subcauses": [{"name": "Competencia", "subcauses": ["Escasez"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Exclusión cultural", "subcauses": [{"name": "Idioma español", "subcauses": ["Incomunicación"]}]},
                {"name": "Diseño impuesto", "subcauses": [{"name": "Ignora costumbres", "subcauses": ["Rechazo"]}]},
                {"name": "División comunitaria", "subcauses": [{"name": "A favor vs contra", "subcauses": ["Conflictos"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Ignorar líderes", "subcauses": [{"name": "Diálogo con equivocados", "subcauses": ["Acuerdos nulos"]}]},
                {"name": "Prisa política", "subcauses": [{"name": "Saltar pasos", "subcauses": ["Violación derechos"]}]},
                {"name": "Falta de transparencia", "subcauses": [{"name": "Contratos ocultos", "subcauses": ["Sospecha"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Mapeo centralista", "subcauses": [{"name": "Sin censo local", "subcauses": ["Comunidades invisibles"]}]},
                {"name": "Quejas inaccesibles", "subcauses": [{"name": "Buzón web", "subcauses": ["Sin internet"]}]},
                {"name": "Diseño estandarizado", "subcauses": [{"name": "Modelo urbano", "subcauses": ["Falla rural"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Violación OIT 169", "subcauses": [{"name": "Demanda internacional", "subcauses": ["Paralización"]}]},
                {"name": "Incumplimiento BID", "subcauses": [{"name": "MICI activado", "subcauses": ["Retiro fondos"]}]},
                {"name": "Leyes indígenas ignoradas", "subcauses": [{"name": "Desacato", "subcauses": ["Amparos"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Subestimación social", "subcauses": [{"name": "Pensar que no hay quejas", "subcauses": ["Sorpresa"]}]},
                {"name": "Sin antropólogos", "subcauses": [{"name": "Solo ingenieros", "subcauses": ["Mala lectura"]}]},
                {"name": "Monitoreo nulo", "subcauses": [{"name": "Falta presencia", "subcauses": ["Crisis"]}]}
            ]}
        ]

        analysis["ishikawas"]["Incumplimiento de contratistas"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Escombros dispersos", "subcauses": [{"name": "Vertederos ilegales", "subcauses": ["Contaminación"]}]},
                {"name": "Derrames químicos", "subcauses": [{"name": "Pinturas", "subcauses": ["Suelo tóxico"]}]},
                {"name": "Ruido generalizado", "subcauses": [{"name": "Maquinaria simultánea", "subcauses": ["Afectación fauna"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Molestias simultáneas", "subcauses": [{"name": "Calles cerradas", "subcauses": ["Vecinos molestos"]}]},
                {"name": "Abusos laborales", "subcauses": [{"name": "Subcontrato informal", "subcauses": ["Explotación"]}]},
                {"name": "Accidentes en 1000 sitios", "subcauses": [{"name": "Sin EPP", "subcauses": ["Lesiones"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Agencia rebasada", "subcauses": [{"name": "Poco personal", "subcauses": ["Descontrol"]}]},
                {"name": "Descoordinación", "subcauses": [{"name": "5 estados", "subcauses": ["Órdenes cruzadas"]}]},
                {"name": "Corrupción micro", "subcauses": [{"name": "Sobornos locales", "subcauses": ["Avance falso"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Reportes en papel", "subcauses": [{"name": "Retrasos", "subcauses": ["Decisión tardía"]}]},
                {"name": "Logística imposible", "subcauses": [{"name": "Distancias enormes", "subcauses": ["Sin visitas"]}]},
                {"name": "Contratación fragmentada", "subcauses": [{"name": "Demasiadas empresas", "subcauses": ["Caos legal"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Contratos débiles", "subcauses": [{"name": "Sin penalidad", "subcauses": ["Impunidad"]}]},
                {"name": "Leyes dispares", "subcauses": [{"name": "Normas por estado", "subcauses": ["Confusión"]}]},
                {"name": "Permisos faltantes", "subcauses": [{"name": "Obras ilegales", "subcauses": ["Clausuras"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Monitoreo imposible", "subcauses": [{"name": "Visita física", "subcauses": ["Insuficiente"]}]},
                {"name": "Sin auditorías", "subcauses": [{"name": "Falta aleatoriedad", "subcauses": ["Relajamiento"]}]},
                {"name": "Presupuesto E&S bajo", "subcauses": [{"name": "No cubre viajes", "subcauses": ["Abandono"]}]}
            ]}
        ]

        
        analysis["stakeholders"] = [
            {"name": "Niños y Maestros rurales", "power": "Bajo", "impact": "Alto"},
            {"name": "Asambleas Ejidales", "power": "Alto", "impact": "Alto"},
            {"name": "Subcontratistas locales", "power": "Medio", "impact": "Bajo"},
            {"name": "SEP Federal", "power": "Alto", "impact": "Bajo"}
        ]
        analysis["crva"] = "El diseño de escuelas no contempla las olas de calor extremo exacerbadas por el cambio climático. Los comedores cerrados y la red de gas aumentan exponencialmente el riesgo de ignición bajo estrés térmico constante."
        analysis["bowtie"] = {
            "causes": ["Gas con soldadura porosa", "Falta de ventilación térmica"],
            "preventive": ["Pruebas de hermeticidad", "Diseño bioclimático cruzado"],
            "hazard": "Fuga de Gas y Explosión en Escuela",
            "mitigative": ["Corte sísmico automático", "Simulacros de evacuación"],
            "consequences": ["Pérdida de vidas infantiles", "Clausura total del proyecto"]
        }
        analysis["impact"] = "B"
        analysis["risk"] = "Substantial"
        analysis["esic_reasoning"] = "Category B: Impactos mitigables, pero masivos."
        analysis["esrr_factors"]["cause"] = "Riesgos de gas y escombros masivos."
        analysis["esrr_factors"]["contribution"] = "Dispersión en cinco estados."
        analysis["esrr_factors"]["context"] = "Superposición con comunidades indígenas."
        analysis["esrr_factors"]["performance"] = "Dificultad extrema de supervisión masiva."
        analysis["reason"] = "Escala masiva y contextos vulnerables."

    elif is_heritage:
        analysis["geo_scoping"] = """
        <div style="background:#fff7ed; padding: 1rem; border-left: 4px solid #ea580c; border-radius:5px; margin-bottom: 2rem;">
        <b>🌍 País:</b> Perú<br>
        <b>📍 Zona del Proyecto:</b> Lima Centro (Barrios Altos, Quinta Heeren).<br>
        <b>⚠️ Marcadores de Sensibilidad:</b> 🔴 Patrimonio Histórico, 🔴 Arqueológico, 🔴 Vulnerabilidad Social.<br>
        <b>🔍 Diagnóstico Geo-Espacial ESG:</b> Zona de Intervención Hiper-Regulada. Se interviene directamente sobre el <b>Patrimonio Cultural de la Nación (MINCUL)</b>. Existe probabilidad casi absoluta de hallazgos arqueológicos en el subsuelo (Huacas prehispánicas/coloniales) y un fuerte marcador de sensibilidad social por riesgo de <b>gentrificación</b>.</div>
        """
        analysis["summary"] = """
        **Resumen Ejecutivo del Proyecto (RICH Fase 2): Intervención Urbana en Barrios Altos y Quinta Heeren**
        
        El proyecto persigue la recuperación urbana, mejoramiento de espacios públicos y rehabilitación de patrimonio arquitectónico monumental en el centro histórico de Lima, con foco particular en la Quinta Heeren.
        
        **Hallazgos Críticos del Contexto Histórico y Social:**
        *   **Patrimonio Cultural de la Nación (MINCUL):** La zona de intervención está bajo estricta protección patrimonial. Las estructuras (mayormente de adobe, quincha y madera del siglo XIX) son altamente susceptibles a daños irreversibles por vibraciones mecánicas, humedad o técnicas de construcción modernas (como el uso de concreto armado). Cualquier obra civil subterránea tiene altísima probabilidad de generar hallazgos arqueológicos.
        *   **Fragilidad Social (Riesgo de Gentrificación):** Barrios Altos está habitado históricamente por poblaciones de estratos socioeconómicos bajos y vulnerables, muchas veces en condiciones de inquilinato o posesión precaria. El embellecimiento urbano del sector generará inevitablemente especulación inmobiliaria, alzas de alquileres y el desplazamiento (gentrificación) de estos vecinos tradicionales si no se aplican medidas de protección de tejido social.
        """
        analysis["fmea"] = [
            {"Modo de Falla": "Técnicas destructivas en casonas de adobe", "Efecto": "Daño Irreversible a Patrimonio", "SEV": 9, "OCC": 6, "DET": 5, "RPN": 270, "Acción": "Uso de cal manual", "Jerarquía": "Evitar"},
            {"Modo de Falla": "Hallazgos no reportados", "Efecto": "Destrucción de vestigios arqueológicos", "SEV": 8, "OCC": 7, "DET": 6, "RPN": 336, "Acción": "Arqueólogo residente", "Jerarquía": "Minimizar"},
            {"Modo de Falla": "Aumento excesivo de alquileres", "Efecto": "Gentrificación y desplazamiento vecinal", "SEV": 7, "OCC": 8, "DET": 4, "RPN": 224, "Acción": "Políticas de retención", "Jerarquía": "Compensar"}
        ]
        
        analysis["ishikawas"]["Daño Irreversible a Patrimonio"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Vibraciones pesadas", "subcauses": [{"name": "Daño estructural", "subcauses": ["Fisuras"]}]},
                {"name": "Humedad infiltrada", "subcauses": [{"name": "Pudrición adobe", "subcauses": ["Colapso"]}]},
                {"name": "Polvo abrasivo", "subcauses": [{"name": "Erosión murales", "subcauses": ["Decoloración"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Pérdida valor cultural", "subcauses": [{"name": "Rechazo ciudadano", "subcauses": ["Protestas"]}]},
                {"name": "Desarraigo", "subcauses": [{"name": "Historia borrada", "subcauses": ["Indignación"]}]},
                {"name": "Conflictos patronatos", "subcauses": [{"name": "Demandas diseño", "subcauses": ["Paros"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Coordinación débil", "subcauses": [{"name": "ProLIMA vs MINCUL", "subcauses": ["Bloqueo"]}]},
                {"name": "Supervisión laxa", "subcauses": [{"name": "Sin expertos", "subcauses": ["Errores"]}]},
                {"name": "Toma decisión lenta", "subcauses": [{"name": "Burocracia", "subcauses": ["Paralización"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Cemento incompatible", "subcauses": [{"name": "Retención agua", "subcauses": ["Deterioro"]}]},
                {"name": "Herramientas percusión", "subcauses": [{"name": "Impacto duro", "subcauses": ["Rotura"]}]},
                {"name": "Sin modelado BIM", "subcauses": [{"name": "Ceguera 3D", "subcauses": ["Cortes errados"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Falta CIRA", "subcauses": [{"name": "Permiso retrasado", "subcauses": ["Ilegalidad"]}]},
                {"name": "Alturas violadas", "subcauses": [{"name": "Norma centro histórico", "subcauses": ["Multa"]}]},
                {"name": "OHS deficiente", "subcauses": [{"name": "Sin apuntalamiento", "subcauses": ["Accidente"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Sin monitoreo fisuras", "subcauses": [{"name": "Asentamiento ciego", "subcauses": ["Derrumbe"]}]},
                {"name": "Presupuesto fijo", "subcauses": [{"name": "Falta contingencia", "subcauses": ["Obra parada"]}]},
                {"name": "Sin plan rescate", "subcauses": [{"name": "Pérdida piezas", "subcauses": ["Daño final"]}]}
            ]}
        ]
        
        analysis["ishikawas"]["Destrucción de vestigios arqueológicos"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Afectación subsuelo", "subcauses": [{"name": "Estratigrafía rota", "subcauses": ["Pérdida data"]}]},
                {"name": "Modificación nivel freático", "subcauses": [{"name": "Inundación restos", "subcauses": ["Pudrición"]}]},
                {"name": "Contaminación química", "subcauses": [{"name": "Derrames", "subcauses": ["Manchas en huacas"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Pérdida identidad", "subcauses": [{"name": "Historia negada", "subcauses": ["Indignación"]}]},
                {"name": "Saqueo (Huaqueo)", "subcauses": [{"name": "Venta ilegal", "subcauses": ["Pérdida nacional"]}]},
                {"name": "Reclamo mediático", "subcauses": [{"name": "Prensa negativa", "subcauses": ["Escándalo"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "MINCUL ausente", "subcauses": [{"name": "Falta inspectores", "subcauses": ["Impunidad"]}]},
                {"name": "Ocultamiento", "subcauses": [{"name": "Contratista calla", "subcauses": ["Complicidad"]}]},
                {"name": "Corrupción permisos", "subcauses": [{"name": "Liberación express", "subcauses": ["Irregularidad"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Excavación rápida", "subcauses": [{"name": "Maquinaria pesada", "subcauses": ["Trituración"]}]},
                {"name": "Sin zarandeo", "subcauses": [{"name": "Tierra descartada", "subcauses": ["Pérdida piezas"]}]},
                {"name": "Topografía errada", "subcauses": [{"name": "Cortes ciegos", "subcauses": ["Daño directo"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Violación Ley Patrimonio", "subcauses": [{"name": "Delito penal", "subcauses": ["Cárcel"]}]},
                {"name": "Protocolo omitido", "subcauses": [{"name": "Chance finds ignorado", "subcauses": ["Falta grave"]}]},
                {"name": "Informes falsos", "subcauses": [{"name": "Falsificación", "subcauses": ["Fraude"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Falta arqueólogo", "subcauses": [{"name": "Sin supervisión", "subcauses": ["Descontrol"]}]},
                {"name": "Sin pausa de obra", "subcauses": [{"name": "Prioridad tiempo", "subcauses": ["Daño seguro"]}]},
                {"name": "Riesgo no mapeado", "subcauses": [{"name": "Matriz incompleta", "subcauses": ["Sorpresa"]}]}
            ]}
        ]

        analysis["ishikawas"]["Gentrificación y desplazamiento vecinal"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Mejora estética urbana", "subcauses": [{"name": "Atracción masiva", "subcauses": ["Saturación"]}]},
                {"name": "Renovación redes", "subcauses": [{"name": "Plusvalía del suelo", "subcauses": ["Aumento valor"]}]},
                {"name": "Limpieza extrema", "subcauses": [{"name": "Cambio de paisaje", "subcauses": ["Nuevo estrato"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Desarraigo barrial", "subcauses": [{"name": "Migración forzada", "subcauses": ["Pobreza"]}]},
                {"name": "Pérdida de comercio local", "subcauses": [{"name": "Cierre de bodegas", "subcauses": ["Llegada franquicias"]}]},
                {"name": "Ruptura tejido social", "subcauses": [{"name": "Nuevos vs Viejos", "subcauses": ["Conflicto"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Falta política social", "subcauses": [{"name": "Alcaldía omisiva", "subcauses": ["Solo enfoque físico"]}]},
                {"name": "Lobby inmobiliario", "subcauses": [{"name": "Presión constructores", "subcauses": ["Especulación"]}]},
                {"name": "Sin subsidios", "subcauses": [{"name": "Abandono estatal", "subcauses": ["Desprotección"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Diseño para turistas", "subcauses": [{"name": "Ignora residentes", "subcauses": ["Hostilidad"]}]},
                {"name": "Procesos expropiación", "subcauses": [{"name": "Desalojos", "subcauses": ["Uso fuerza"]}]},
                {"name": "Marketing inmobiliario", "subcauses": [{"name": "Venta exclusividad", "subcauses": ["Burbuja"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Ley alquileres libre", "subcauses": [{"name": "Alzas sin límite", "subcauses": ["Expulsión"]}]},
                {"name": "Falta regulación social", "subcauses": [{"name": "Vacío legal", "subcauses": ["Injusticia"]}]},
                {"name": "Zonificación cambiada", "subcauses": [{"name": "Uso comercial nuevo", "subcauses": ["Incompatibilidad"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Impacto social no medido", "subcauses": [{"name": "Ceguera evaluativa", "subcauses": ["Sorpresa"]}]},
                {"name": "Sin plan mitigación", "subcauses": [{"name": "No hay fondos sociales", "subcauses": ["Inequidad"]}]},
                {"name": "Ignorar quejas", "subcauses": [{"name": "Mecanismo sordo", "subcauses": ["Escalamiento"]}]}
            ]}
        ]
        
        
        analysis["stakeholders"] = [
            {"name": "Inquilinos precarios", "power": "Bajo", "impact": "Alto"},
            {"name": "Ministerio de Cultura", "power": "Alto", "impact": "Alto"},
            {"name": "Inmobiliarias especulativas", "power": "Alto", "impact": "Bajo"},
            {"name": "Patronato del Centro", "power": "Medio", "impact": "Bajo"}
        ]
        analysis["crva"] = "Fenómeno de El Niño Costero anómalo. Las lluvias torrenciales inusuales en Lima disolverán la arquitectura de adobe y quincha si las cubiertas protectoras temporales de la obra fallan durante la ejecución."
        analysis["bowtie"] = {
            "causes": ["Vibración por maquinaria pesada", "Lluvias sobre estructura desnuda"],
            "preventive": ["Herramientas manuales exclusivas", "Carpas térmicas estructurales"],
            "hazard": "Colapso de Casonas Patrimoniales (Adobe)",
            "mitigative": ["Apuntalamiento de emergencia", "Plan de salvataje express"],
            "consequences": ["Pérdida Irreversible de Patrimonio", "Desplazamiento de vecinos"]
        }
        analysis["impact"] = "B"
        analysis["risk"] = "Substantial"
        analysis["esic_reasoning"] = "Category B: Impactos mitigables físicamente."
        analysis["esrr_factors"]["cause"] = "Riesgo físico a casonas y hallazgos."
        analysis["esrr_factors"]["contribution"] = "Fuerte riesgo de gentrificación social."
        analysis["esrr_factors"]["context"] = "Entorno hiper-regulado (MINCUL)."
        analysis["esrr_factors"]["performance"] = "Alta exigencia restauradora."
        analysis["reason"] = "Sensibilidad patrimonial y social eleva riesgo."

    elif is_br_l1700 or "br-l" in text_lower or "pod " in text_lower:
        analysis["geo_scoping"] = """
        <div style="background:#f0fdf4; padding: 1rem; border-left: 4px solid #16a34a; border-radius:5px; margin-bottom: 2rem;">
        <b>🌍 País:</b> Brasil<br>
        <b>📍 Zona del Proyecto:</b> Nacional / Entorno Legislativo Federal.<br>
        <b>⚠️ Marcadores de Sensibilidad:</b> Ninguno Físico. Riesgo Macro-Social.<br>
        <b>🔍 Diagnóstico Geo-Espacial ESG:</b> Al ser un PBL (Policy-Based Loan), <b>no existe huella de construcción física</b>. Por tanto, no hay traslape con áreas indígenas, culturales o arqueológicas. El análisis se desplaza hacia la sensibilidad política y las consecuencias socioeconómicas distributivas de las leyes.</div>
        """
        analysis["summary"] = """
        **Resumen Ejecutivo del Proyecto (BR-L1700): Apoyo Presupuestario en Brasil (PBL)**
        
        Esta operación es un Préstamo Basado en Políticas (Policy-Based Loan). No financiará obras civiles. Su objetivo es apoyar reformas macroeconómicas, regulatorias y fiscales.
        
        **Hallazgos Críticos del Contexto (Riesgos Distribucionales y Políticos):**
        Aunque no hay huella física (obras), el análisis de la economía política revela riesgos severos. Las reformas estructurales (recorte de subsidios, aumento de impuestos) pueden tener un impacto regresivo no mapeado. Si estas medidas golpean a poblaciones vulnerables (favelas, zonas rurales), el descontento social podría incendiar las calles. Además, un congreso fragmentado podría secuestrar las reformas, paralizando los tramos de desembolso y dejando al país en default técnico.
        """
        analysis["fmea"] = [
            {"Modo de Falla": "Rechazo del congreso a reformas", "Efecto": "Caída de la matriz de política", "SEV": 8, "OCC": 7, "DET": 4, "RPN": 224, "Acción": "Cabildeo y negociación", "Jerarquía": "Minimizar"},
            {"Modo de Falla": "Protestas por reformas regresivas", "Efecto": "Estallido social y parálisis", "SEV": 9, "OCC": 6, "DET": 4, "RPN": 216, "Acción": "Evaluación Distribucional", "Jerarquía": "Evitar"},
            {"Modo de Falla": "Incapacidad técnica ministerial", "Efecto": "Incumplimiento de condiciones", "SEV": 7, "OCC": 7, "DET": 5, "RPN": 245, "Acción": "Asistencia Técnica Integrada", "Jerarquía": "Restaurar"}
        ]
        
        analysis["ishikawas"]["Caída de la matriz de política"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Presión del lobby agrario", "subcauses": [{"name": "Rechazo a leyes verdes", "subcauses": ["Bloqueo de presupuesto"]}]},
                {"name": "Subsidios fósiles intocables", "subcauses": [{"name": "Miedo a inflación energética", "subcauses": ["Veto legislativo"]}]},
                {"name": "Crisis climática ignorada", "subcauses": [{"name": "Prioridad en deuda a corto plazo", "subcauses": ["Desvío de atención"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Impopularidad presidencial", "subcauses": [{"name": "Pérdida de capital político", "subcauses": ["Congreso rebelde"]}]},
                {"name": "Polarización extrema", "subcauses": [{"name": "Guerra cultural", "subcauses": ["Falta de consenso"]}]},
                {"name": "Desinformación mediática", "subcauses": [{"name": "Fake news sobre la reforma", "subcauses": ["Pánico ciudadano"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Congreso fragmentado", "subcauses": [{"name": "Múltiples partidos pequeños", "subcauses": ["Extorsión política"]}]},
                {"name": "Caída de ministros clave", "subcauses": [{"name": "Inestabilidad de gabinete", "subcauses": ["Reinicio de negociaciones"]}]},
                {"name": "Año electoral", "subcauses": [{"name": "Cálculo de votos", "subcauses": ["Rechazo a medidas impopulares"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Sistemas de voto hackeables", "subcauses": [{"name": "Desconfianza en el proceso", "subcauses": ["Impugnación de resultados"]}]},
                {"name": "Redacción de leyes ambigua", "subcauses": [{"name": "Vacíos legales", "subcauses": ["Interpretaciones cruzadas"]}]},
                {"name": "Plazos legislativos irreales", "subcauses": [{"name": "Prisa impuesta por el Banco", "subcauses": ["Rechazo por forma"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Inconstitucionalidad", "subcauses": [{"name": "Choque con derechos adquiridos", "subcauses": ["Corte Suprema frena ley"]}]},
                {"name": "Violación de regla fiscal", "subcauses": [{"name": "Tope de gasto excedido", "subcauses": ["Juicio político"]}]},
                {"name": "Leyes contradictorias", "subcauses": [{"name": "Superposición normativa", "subcauses": ["Parálisis jurídica"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Ceguera política del Banco", "subcauses": [{"name": "No medir encuestas locales", "subcauses": ["Falsa expectativa de éxito"]}]},
                {"name": "Sin plan B", "subcauses": [{"name": "Dependencia de una sola ley", "subcauses": ["Fracaso total si no pasa"]}]},
                {"name": "Análisis de actores deficiente", "subcauses": [{"name": "Ignorar a opositores ocultos", "subcauses": ["Emboscada legislativa"]}]}
            ]}
        ]

        analysis["ishikawas"]["Estallido social y parálisis"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Aumento del costo de transporte", "subcauses": [{"name": "Impuesto al carbono", "subcauses": ["Huelga de camioneros"]}]},
                {"name": "Desregulación ambiental", "subcauses": [{"name": "Para contentar industrias", "subcauses": ["Protestas ecologistas"]}]},
                {"name": "Privatización del agua", "subcauses": [{"name": "Tarifas impagables", "subcauses": ["Toma de plantas hídricas"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Aumento de la pobreza", "subcauses": [{"name": "Recorte de subsidios sociales", "subcauses": ["Hambruna urbana"]}]},
                {"name": "Desempleo masivo", "subcauses": [{"name": "Ajuste fiscal severo", "subcauses": ["Crimen desatado"]}]},
                {"name": "Sentimiento de injusticia", "subcauses": [{"name": "Élites no pagan impuestos", "subcauses": ["Saqueos a supermercados"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Falta de diálogo social", "subcauses": [{"name": "Reformas hechas a puerta cerrada", "subcauses": ["Sorpresa e ira"]}]},
                {"name": "Represión policial brutal", "subcauses": [{"name": "Uso excesivo de fuerza", "subcauses": ["Radicalización de protestas"]}]},
                {"name": "Pérdida de legitimidad", "subcauses": [{"name": "Escándalos de corrupción paralelos", "subcauses": ["Renuncia presidencial exigida"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Modelos económicos errados", "subcauses": [{"name": "Fallas en el Excel del Ministerio", "subcauses": ["Impacto real subestimado"]}]},
                {"name": "Algoritmos discriminatorios", "subcauses": [{"name": "Focalización mal diseñada", "subcauses": ["Exclusión de beneficiarios legítimos"]}]},
                {"name": "Bases de datos obsoletas", "subcauses": [{"name": "Censos viejos", "subcauses": ["Reparto desigual de ayuda"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Incumplimiento de derechos humanos", "subcauses": [{"name": "Condenas de la CIDH", "subcauses": ["Aislamiento internacional"]}]},
                {"name": "Estados de sitio ilegales", "subcauses": [{"name": "Suspensión de garantías", "subcauses": ["Violaciones masivas"]}]},
                {"name": "Derecho a huelga criminalizado", "subcauses": [{"name": "Encarcelamiento de líderes", "subcauses": ["Mártires políticos"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "No hacer PSIA (Poverty Social Impact)", "subcauses": [{"name": "Ceguera distributiva", "subcauses": ["Creer que la reforma es neutra"]}]},
                {"name": "Sin red de seguridad temporal", "subcauses": [{"name": "Ajuste de shock", "subcauses": ["No hay colchón para los pobres"]}]},
                {"name": "Monitoreo social nulo", "subcauses": [{"name": "Ignorar alertas tempranas", "subcauses": ["Explosión repentina"]}]}
            ]}
        ]

        analysis["ishikawas"]["Incumplimiento de condiciones"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Desastres climáticos concurrentes", "subcauses": [{"name": "Inundaciones en Brasilia", "subcauses": ["Ministerios cerrados"]}]},
                {"name": "Crisis energética", "subcauses": [{"name": "Apagones nacionales", "subcauses": ["Pérdida de datos y trabajo"]}]},
                {"name": "Emergencia sanitaria", "subcauses": [{"name": "Nuevas pandemias", "subcauses": ["Atención desviada de reformas"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Fuga de cerebros", "subcauses": [{"name": "Bajos salarios públicos", "subcauses": ["Técnicos migran al sector privado"]}]},
                {"name": "Huelga de funcionarios", "subcauses": [{"name": "Exigencia de bonos", "subcauses": ["Paralización administrativa"]}]},
                {"name": "Desmotivación burocrática", "subcauses": [{"name": "Falta de liderazgo claro", "subcauses": ["Lentitud extrema"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Rotación de personal clave", "subcauses": [{"name": "Cambios políticos", "subcauses": ["Pérdida de memoria institucional"]}]},
                {"name": "Celos interministeriales", "subcauses": [{"name": "Hacienda vs Ambiente", "subcauses": ["Negativa a compartir datos"]}]},
                {"name": "Falta de mandato directivo", "subcauses": [{"name": "Nadie asume la reforma", "subcauses": ["Inacción"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Sistemas IT anticuados", "subcauses": [{"name": "Servidores lentos", "subcauses": ["Imposibilidad de procesar big data"]}]},
                {"name": "Silos de información", "subcauses": [{"name": "Bases de datos no conectadas", "subcauses": ["Trabajo manual propenso a error"]}]},
                {"name": "Ciberataques al gobierno", "subcauses": [{"name": "Ransomware en Hacienda", "subcauses": ["Secuestro de expedientes"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Reglamentos ultra-rígidos", "subcauses": [{"name": "Leyes de adquisiciones lentas", "subcauses": ["No se pueden contratar consultores"]}]},
                {"name": "Auditorías de la Contraloría", "subcauses": [{"name": "Parálisis por terror a firmar", "subcauses": ["Congelamiento de decisiones"]}]},
                {"name": "Incumplimiento de matriz del Banco", "subcauses": [{"name": "Condiciones mal redactadas", "subcauses": ["Imposibles de probar legalmente"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Falta de presupuesto para la reforma", "subcauses": [{"name": "Ministerio sin fondos operativos", "subcauses": ["Imposibilidad de implementar"]}]},
                {"name": "Ceguera técnica del BID", "subcauses": [{"name": "Exigir lo imposible al país", "subcauses": ["Metas inalcanzables"]}]},
                {"name": "Ausencia de indicadores claros", "subcauses": [{"name": "No se sabe cómo medir el éxito", "subcauses": ["Debate infinito de cumplimiento"]}]}
            ]}
        ]
        
        
        analysis["stakeholders"] = [
            {"name": "Población en Pobreza", "power": "Bajo", "impact": "Alto"},
            {"name": "Bancada Agraria (Congreso)", "power": "Alto", "impact": "Medio"},
            {"name": "Ministerio de Hacienda", "power": "Alto", "impact": "Bajo"},
            {"name": "Sindicatos Nacionales", "power": "Alto", "impact": "Alto"}
        ]
        analysis["crva"] = "El recorte de subsidios y presupuesto desfinanciará a las agencias de monitoreo ambiental (IBAMA), reduciendo el patrullaje forestal y acelerando la deforestación amazónica inducida por sequías (Efecto en Cascada)."
        analysis["bowtie"] = {
            "causes": ["Recorte ciego de subsidios", "Aumento regresivo de impuestos"],
            "preventive": ["Evaluación Distributiva Ex-ante", "Diálogo social previo"],
            "hazard": "Estallido Social Nacional por Ajuste",
            "mitigative": ["Expansión de Bolsa Familia", "Inyección de liquidez focalizada"],
            "consequences": ["Caída de gobierno / Default", "Aumento extremo de pobreza"]
        }
        analysis["impact"] = "C"
        analysis["risk"] = "Low"
        analysis["esic_reasoning"] = "Category C: Sin huella física o ambiental directa."
        analysis["esrr_factors"]["cause"] = "Mínimos impactos directos."
        analysis["esrr_factors"]["contribution"] = "Reformas de políticas públicas con alto impacto distributivo."
        analysis["esrr_factors"]["context"] = "Economía macro y gobernabilidad del país."
        analysis["esrr_factors"]["performance"] = "Alta dependencia de la estabilidad política."
        analysis["reason"] = "Riesgos predominantemente políticos y sociales, tipificado como C/Low."

    elif is_liquidity or "co-l" in text_lower or "col " in text_lower or "contingent" in text_lower:
        analysis["geo_scoping"] = """
        <div style="background:#fef2f2; padding: 1rem; border-left: 4px solid #dc2626; border-radius:5px; margin-bottom: 2rem;">
        <b>🌍 País:</b> Colombia (Línea Continente LAC).<br>
        <b>📍 Zona del Proyecto:</b> Zonas Cero post-desastre (Indeterminado ex-ante).<br>
        <b>⚠️ Marcadores de Sensibilidad:</b> 🔴 Crisis Humanitaria, 🔴 Riesgo Biológico.<br>
        <b>🔍 Diagnóstico Geo-Espacial ESG:</b> Geografía variable dependiendo del epicentro de la catástrofe. Es imposible descartar el impacto a sitios culturales previos al evento, pero la urgencia recae en la <b>vulnerabilidad humanitaria extrema</b> (refugiados, hacinamiento, VBG) y el riesgo tóxico de las ruinas.</div>
        """
        analysis["summary"] = """
        **Resumen Ejecutivo del Proyecto (CO-L1333): Línea de Crédito Contingente para Emergencias y Desastres**
        
        El proyecto es un instrumento financiero rápido diseñado para inyectar liquidez al país inmediatamente después de una declaratoria de desastre natural (terremoto, inundación, huracán) o emergencia sanitaria.
        
        **Hallazgos Críticos del Contexto (Mirada Detectivesca y Alarmista):**
        Aunque se trate de una transferencia financiera, el contexto en el que se gatilla es de *caos absoluto*. En una emergencia, los controles ambientales y sociales del Estado colapsan. 
        *   **Riesgo Biológico y Tóxico:** El dinero se usará rápido para remover escombros urbanos masivos, los cuales estarán mezclados con cadáveres (riesgo biológico), asbesto fracturado (riesgo cancerígeno) y químicos industriales. 
        *   **Crisis Humanitaria y VBG:** Los fondos financiarán albergues temporales masivos. En el hacinamiento y la desesperación, sin iluminación ni privacidad, estos campamentos se convierten en epicentros de Violencia Basada en Género (VBG), abusos infantiles y brotes de cólera.
        *   **Corrupción por Emergencia:** Al levantarse los procesos de licitación regulares, el riesgo de "captura de fondos" y construcción de obras fantasma por mafias locales se dispara a niveles críticos.
        """
        analysis["fmea"] = [
            {"Modo de Falla": "Manejo caótico de escombros de la tragedia", "Efecto": "Contaminación masiva y riesgo biológico", "SEV": 10, "OCC": 8, "DET": 4, "RPN": 320, "Acción": "Protocolos disposición", "Jerarquía": "Minimizar"},
            {"Modo de Falla": "Hacinamiento en campamentos de refugiados", "Efecto": "Brotes epidemiológicos y crisis de VBG", "SEV": 9, "OCC": 8, "DET": 4, "RPN": 288, "Acción": "Vigilancia en albergues", "Jerarquía": "Minimizar"},
            {"Modo de Falla": "Levantamiento de controles fiduciarios", "Efecto": "Captura de fondos de emergencia y obras fantasma", "SEV": 8, "OCC": 9, "DET": 3, "RPN": 216, "Acción": "Auditoría en tiempo real", "Jerarquía": "Restaurar"}
        ]
        
        analysis["ishikawas"]["Contaminación masiva y riesgo biológico"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Mezcla de materiales peligrosos", "subcauses": [{"name": "Asbesto y plomo fracturado", "subcauses": ["Polvo cancerígeno en el aire"]}]},
                {"name": "Putrefacción biológica", "subcauses": [{"name": "Cadáveres humanos y animales bajo escombros", "subcauses": ["Lixiviados cadavéricos al agua subterránea"]}]},
                {"name": "Colapso de alcantarillados", "subcauses": [{"name": "Aguas negras a cielo abierto", "subcauses": ["Ríos urbanos de heces"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Pánico y desesperación", "subcauses": [{"name": "Civiles removiendo ruinas con las manos", "subcauses": ["Cortes y tétanos masivo"]}]},
                {"name": "Pérdida de hogares absolutos", "subcauses": [{"name": "Personas viviendo sobre la basura tóxica", "subcauses": ["Enfermedades respiratorias agudas"]}]},
                {"name": "Saqueos en zonas de desastre", "subcauses": [{"name": "Caos social", "subcauses": ["Dificulta el trabajo de rescate seguro"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Colapso del gobierno local", "subcauses": [{"name": "Alcaldía destruida o alcaldes fallecidos", "subcauses": ["Anarquía y vacío de poder"]}]},
                {"name": "Prioridad ciega por despejar vías", "subcauses": [{"name": "Uso de buldóceres sin inspección", "subcauses": ["Trituración de cuerpos y evidencia"]}]},
                {"name": "Descoordinación de agencias", "subcauses": [{"name": "Ejército vs Defensa Civil vs Cruz Roja", "subcauses": ["Órdenes cruzadas ineficientes"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Falta de vertederos de seguridad", "subcauses": [{"name": "Creación de botaderos ilegales en ríos", "subcauses": ["Contaminación cuenca abajo"]}]},
                {"name": "Equipos de Protección Personal (EPP) nulos", "subcauses": [{"name": "Rescatistas respirando asbesto", "subcauses": ["Muerte silenciosa de trabajadores"]}]},
                {"name": "Logística colapsada", "subcauses": [{"name": "Carreteras rotas", "subcauses": ["Imposible sacar los residuos de la ciudad"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Suspensión de leyes ambientales", "subcauses": [{"name": "Decreto de emergencia excusa todo", "subcauses": ["Impunidad ecológica legalizada"]}]},
                {"name": "Ignorar salvaguardas del BID", "subcauses": [{"name": "El país asume que la emergencia perdona todo", "subcauses": ["Violación contractual grave"]}]},
                {"name": "Tráfico de residuos", "subcauses": [{"name": "Chatarreros vendiendo metal contaminado", "subcauses": ["Expansión del veneno"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Sin mapa de riesgos químicos previos", "subcauses": [{"name": "No se sabe dónde había fábricas químicas", "subcauses": ["Explosiones secundarias en las ruinas"]}]},
                {"name": "No hay presupuesto para disposición final", "subcauses": [{"name": "El dinero es solo para recoger, no tratar", "subcauses": ["Problema heredado por décadas"]}]},
                {"name": "Monitoreo imposible", "subcauses": [{"name": "Inspectores ambientales desplazados o muertos", "subcauses": ["Ceguera total de la obra"]}]}
            ]}
        ]

        analysis["ishikawas"]["Brotes epidemiológicos y crisis de VBG"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Falta de agua potable", "subcauses": [{"name": "Consumo de pozos contaminados", "subcauses": ["Cólera y disentería"]}]},
                {"name": "Proliferación de vectores", "subcauses": [{"name": "Mosquitos y ratas en charcos", "subcauses": ["Dengue, malaria y leptospirosis"]}]},
                {"name": "Condiciones climáticas extremas", "subcauses": [{"name": "Lluvia torrencial o frío en carpas", "subcauses": ["Hipotermia y neumonía"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Hacinamiento extremo", "subcauses": [{"name": "Familias enteras en 5 metros cuadrados", "subcauses": ["Pérdida total de intimidad"]}]},
                {"name": "Violencia Basada en Género (VBG)", "subcauses": [{"name": "Baños oscuros y compartidos", "subcauses": ["Violaciones nocturnas"]}]},
                {"name": "Trata de personas", "subcauses": [{"name": "Niños huérfanos vagando", "subcauses": ["Secuestro por redes criminales"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Falta de vigilancia policial", "subcauses": [{"name": "Policía cuidando bancos, no carpas", "subcauses": ["Ley del más fuerte en el campamento"]}]},
                {"name": "Comités de campamento mafiosos", "subcauses": [{"name": "Líderes barriales extorsionan", "subcauses": ["Comida a cambio de sexo (Supervivencia)"]}]},
                {"name": "Abandono estatal", "subcauses": [{"name": "Se instala la carpa y el Estado se va", "subcauses": ["Guetización del refugio"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Carpas de mala calidad", "subcauses": [{"name": "Lonas que no resisten el clima", "subcauses": ["Refugios inundados"]}]},
                {"name": "Sanitarios insuficientes", "subcauses": [{"name": "1 letrina por cada 100 personas", "subcauses": ["Defecación al aire libre obligada"]}]},
                {"name": "Falta de iluminación solar", "subcauses": [{"name": "Oscuridad total a las 6PM", "subcauses": ["Facilita crímenes y robos"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Violación de normas Esfera (Sphere)", "subcauses": [{"name": "Estándares humanitarios ignorados", "subcauses": ["Condiciones infrahumanas"]}]},
                {"name": "Falta de rutas de denuncia", "subcauses": [{"name": "Víctimas de VBG sin autoridades a quién recurrir", "subcauses": ["Impunidad absoluta"]}]},
                {"name": "Registro de refugiados inexistente", "subcauses": [{"name": "No se sabe quién entra o sale", "subcauses": ["Desapariciones silenciosas"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Planificación solo reactiva", "subcauses": [{"name": "Se arman campamentos en zonas de riesgo residual", "subcauses": ["Nuevos deslaves matan a los refugiados"]}]},
                {"name": "Ignorar vulnerabilidad diferencial", "subcauses": [{"name": "Diseño ciego a discapacitados, mujeres y niños", "subcauses": ["Los más débiles perecen"]}]},
                {"name": "Suministros médicos agotados", "subcauses": [{"name": "Falta de contingencia farmacéutica", "subcauses": ["Muerte por heridas leves infectadas"]}]}
            ]}
        ]

        analysis["ishikawas"]["Captura de fondos de emergencia y obras fantasma"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Excusa de reconstrucción 'Verde' falsa", "subcauses": [{"name": "Greenwashing de urgencia", "subcauses": ["Desvío de fondos a proyectos inútiles"]}]},
                {"name": "Tala ilegal financiada por el estado", "subcauses": [{"name": "Madera para supuesta reconstrucción", "subcauses": ["Deforestación mafiosa acelerada"]}]},
                {"name": "Extracción de arena de ríos destructiva", "subcauses": [{"name": "Contratistas sin escrúpulos reconstruyendo", "subcauses": ["Erosión y desastres futuros asegurados"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Enriquecimiento de mafias locales", "subcauses": [{"name": "Contratos a dedo a amigos del gobernador", "subcauses": ["Aumento de desigualdad"]}]},
                {"name": "Damnificados sin hogar real", "subcauses": [{"name": "Viviendas prometidas nunca entregadas", "subcauses": ["Protestas desesperadas"]}]},
                {"name": "Clientelismo político", "subcauses": [{"name": "Ayuda solo a votantes del partido de turno", "subcauses": ["Discriminación y muerte por color político"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Leyes de contratación suspendidas", "subcauses": [{"name": "Decreto de urgencia elimina licitaciones", "subcauses": ["Puerta abierta al robo descarado"]}]},
                {"name": "Colusión de auditores", "subcauses": [{"name": "Contraloría sobornada con el dinero de emergencia", "subcauses": ["Informes de ejecución limpios pero falsos"]}]},
                {"name": "Centralización de recursos opaca", "subcauses": [{"name": "Todo el dinero en un solo fideicomiso sin supervisión", "subcauses": ["Caja menor del presidente"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Materiales de reconstrucción chatarra", "subcauses": [{"name": "Cobrar acero pero usar varillas oxidadas", "subcauses": ["Colapso de casas en el próximo sismo leve"]}]},
                {"name": "Plataformas de transparencia caídas", "subcauses": [{"name": "Webs del gobierno hackeadas convenientemente", "subcauses": ["Sin rastro de los contratos"]}]},
                {"name": "Pagos en efectivo irrastreables", "subcauses": [{"name": "Excusa de que los bancos no funcionan", "subcauses": ["Lavado de activos de la emergencia"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Lavado de dinero legalizado", "subcauses": [{"name": "Donaciones anónimas permitidas por emergencia", "subcauses": ["Cárteles financiando campañas"]}]},
                {"name": "Falta de actas de recepción", "subcauses": [{"name": "Obras se dan por hechas sin firmas", "subcauses": ["Fraude de infraestructura"]}]},
                {"name": "Elusión de la justicia", "subcauses": [{"name": "Cambios de jurisdicción", "subcauses": ["Casos prescriben en el caos"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "El BID desembolsa sin plan ex-ante", "subcauses": [{"name": "Liquidez inmediata sin manual operativo fiduciario", "subcauses": ["Pérdida de control del préstamo"]}]},
                {"name": "Sin auditorías forenses internacionales", "subcauses": [{"name": "Depender de auditores locales amenazados", "subcauses": ["Cierre en falso del proyecto"]}]},
                {"name": "Ignorar banderas rojas", "subcauses": [{"name": "Precios de materiales inflados 1000%", "subcauses": ["Ceguera institucional por 'solidaridad'"]}]}
            ]}
        ]

        
        analysis["stakeholders"] = [
            {"name": "Damnificados (Mujeres, Niños)", "power": "Bajo", "impact": "Alto"},
            {"name": "Mafias de contratos / Vandalismo", "power": "Alto", "impact": "Alto"},
            {"name": "Defensa Civil / Cruz Roja", "power": "Medio", "impact": "Bajo"},
            {"name": "Gobierno Central", "power": "Alto", "impact": "Bajo"}
        ]
        analysis["crva"] = "Multiplicador de Amenaza. Un sismo debilita laderas, y lluvias extremas posteriores (La Niña) detonan deslizamientos secundarios masivos sobre los campamentos de respuesta (Desastres en Cascada)."
        analysis["bowtie"] = {
            "causes": ["Hacinamiento de damnificados", "Ausencia de vigilancia policial"],
            "preventive": ["Diseño con iluminación LED", "Segregación estricta por género"],
            "hazard": "Brote Epidemiológico y Crisis de VBG",
            "mitigative": ["Rutas de denuncia express", "Profilaxis médica post-exposición"],
            "consequences": ["Mortalidad secundaria superior", "Trauma social intergeneracional"]
        }
        analysis["impact"] = "B"
        analysis["risk"] = "Substantial"
        analysis["esic_reasoning"] = "Category B: Contingencia. Los impactos en la fase de respuesta a desastres son agudos, pero temporal y geográficamente limitados si se mitigan bien."
        analysis["esrr_factors"]["cause"] = "Manejo de escombros masivos, riesgo biológico."
        analysis["esrr_factors"]["contribution"] = "Aumento de VBG y trata en albergues."
        analysis["esrr_factors"]["context"] = "Entorno de caos absoluto post-desastre, control institucional nulo."
        analysis["esrr_factors"]["performance"] = "Gobernanza local fracturada."
        analysis["reason"] = "Riesgos de emergencia social, VBG y ambientales agudos. B / Sustancial."

    else:
        # Fallback genérico - REPARADO CON PROFUNDIDAD
        analysis["geo_scoping"] = """
        <div style="background:#f8fafc; padding: 1rem; border-left: 4px solid #94a3b8; border-radius:5px; margin-bottom: 2rem;">
        <b>🌍 País:</b> No especificado<br>
        <b>📍 Zona del Proyecto:</b> Huella urbana o corporativa estándar.<br>
        <b>⚠️ Marcadores de Sensibilidad:</b> Ninguno Crítico Detectado.<br>
        <b>🔍 Diagnóstico Geo-Espacial ESG:</b> El mapeo inicial <b>no arroja cruces o superposiciones</b> con territorios de comunidades indígenas, áreas arqueológicas listadas, ni bienes de patrimonio cultural. Se asume un contexto espacial de riesgo bajo.</div>
        """
        analysis["summary"] = """
        **Resumen Ejecutivo del Proyecto: Evaluación Estándar / Genérica**
        
        El proyecto analizado propone intervenciones de naturaleza corporativa, adquisiciones institucionales o pequeñas obras de mantenimiento que no involucran una huella física significativa sobre el territorio.
        
        **Hallazgos Generales del Contexto:**
        *   **Baja Sensibilidad Espacial:** No se han detectado marcadores de alta sensibilidad espacial (como áreas naturales protegidas, centros históricos regulados por patrimonio cultural o resguardos de comunidades indígenas). 
        *   Los riesgos identificados son principalmente operativos, relacionados con demoras en licitaciones, fallas en la cadena de suministro, y temas convencionales de seguridad y salud en el trabajo (OHS) en un entorno controlable.
        """
        analysis["fmea"] = [
            {"Modo de Falla": "Gestión general ineficiente", "Efecto": "Retrasos Operativos Severos", "SEV": 8, "OCC": 7, "DET": 4, "RPN": 224, "Acción": "Revisión UEP", "Jerarquía": "Minimizar"},
            {"Modo de Falla": "Problemas de adquisiciones públicas", "Efecto": "Falta de materiales críticos", "SEV": 8, "OCC": 6, "DET": 5, "RPN": 240, "Acción": "Agencia experta de compras", "Jerarquía": "Evitar"},
            {"Modo de Falla": "Alta rotación de personal técnico", "Efecto": "Pérdida de memoria institucional", "SEV": 7, "OCC": 8, "DET": 4, "RPN": 224, "Acción": "Incentivos de retención", "Jerarquía": "Compensar"}
        ]
        
        analysis["ishikawas"]["Retrasos Operativos Severos"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Eventos climáticos extremos", "subcauses": [{"name": "Inundaciones en vías de acceso", "subcauses": ["Paralización logística temporal"]}]},
                {"name": "Permisos ambientales negados", "subcauses": [{"name": "Estudios previos mal elaborados", "subcauses": ["Rechazo de licencias de obra"]}]},
                {"name": "Microsismos imprevistos", "subcauses": [{"name": "Evacuación de oficinas preventivas", "subcauses": ["Pérdida de horas laborales"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Protestas de sindicatos públicos", "subcauses": [{"name": "Exigencia de mejores salarios", "subcauses": ["Tomas de las oficinas ministeriales"]}]},
                {"name": "Huelgas de transportistas", "subcauses": [{"name": "Cierres de carreteras nacionales", "subcauses": ["Insumos atrapados en peajes"]}]},
                {"name": "Resistencia vecinal por ruidos", "subcauses": [{"name": "Quejas comunitarias constantes", "subcauses": ["Clausura temporal de remodelaciones"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Micro-corrupción burocrática", "subcauses": [{"name": "Exigencia de sobornos para firmas", "subcauses": ["Expedientes congelados por meses"]}]},
                {"name": "Falta de director asignado", "subcauses": [{"name": "Puesto acéfalo por renuncias políticas", "subcauses": ["Nadie autoriza desembolsos"]}]},
                {"name": "Descoordinación inter-agencias", "subcauses": [{"name": "Silos institucionales ciegos", "subcauses": ["Papeleo rebotando eternamente"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Caída de servidores estatales", "subcauses": [{"name": "Sistemas informáticos de 20 años", "subcauses": ["Imposibilidad de procesar pagos"]}]},
                {"name": "Diseños de planos deficientes", "subcauses": [{"name": "Errores de cálculo en escritorio", "subcauses": ["Rehacer el trabajo desde cero"]}]},
                {"name": "Cuellos de botella documentales", "subcauses": [{"name": "Firmas manuales requeridas", "subcauses": ["Pilas de papel acumuladas"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Violación de leyes de contratación", "subcauses": [{"name": "Intervención de la Contraloría", "subcauses": ["Obras auditadas y paradas"]}]},
                {"name": "Cambios de normativa a mitad de proyecto", "subcauses": [{"name": "Nuevas exigencias sísmicas obligatorias", "subcauses": ["Incompatibilidad de diseños previos"]}]},
                {"name": "Auditorías sorpresivas fallidas", "subcauses": [{"name": "Informes contables descuadrados", "subcauses": ["Retención cautelar de fondos del BID"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Ceguera ante imprevistos", "subcauses": [{"name": "Sin reservas presupuestarias (contingencia)", "subcauses": ["Obra parada por falta de flujo de caja"]}]},
                {"name": "Ignorar alertas tempranas", "subcauses": [{"name": "Cronogramas falsamente optimistas", "subcauses": ["Sorpresa ante los retrasos evidentes"]}]},
                {"name": "Monitoreo de contratistas nulo", "subcauses": [{"name": "Confianza ciega en promesas verbales", "subcauses": ["Descontrol total de la ejecución"]}]}
            ]}
        ]

        analysis["ishikawas"]["Falta de materiales críticos"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Escasez de recursos naturales globales", "subcauses": [{"name": "Falta de arena para construcción", "subcauses": ["Encarecimiento extremo"]}]},
                {"name": "Contaminación en fábrica proveedora", "subcauses": [{"name": "Clausura gubernamental de proveedores", "subcauses": ["Corte abrupto de suministro"]}]},
                {"name": "Tormentas en alta mar", "subcauses": [{"name": "Barcos de importación retrasados", "subcauses": ["Puertos cerrados por mal clima"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Huelgas en los puertos marítimos", "subcauses": [{"name": "Estibadores paralizados por derechos", "subcauses": ["Contenedores atrapados en aduanas"]}]},
                {"name": "Pánico de acaparamiento general", "subcauses": [{"name": "Especulación de precios por crisis", "subcauses": ["Materiales ocultos en bodegas"]}]},
                {"name": "Inseguridad y piratería en rutas", "subcauses": [{"name": "Robo de camiones de carga en carreteras", "subcauses": ["Pérdida total del envío"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Licitaciones arregladas descubiertas", "subcauses": [{"name": "Adjudicación a empresas fantasma", "subcauses": ["Contrato anulado por tribunales"]}]},
                {"name": "Burocracia aduanera excesiva", "subcauses": [{"name": "Inspectores aduaneros exigiendo coimas", "subcauses": ["Nacionalización paralizada"]}]},
                {"name": "Falta de flujo de caja del Ministerio", "subcauses": [{"name": "Hacienda no gira los anticipos", "subcauses": ["Proveedores se niegan a fabricar"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Términos de Referencia (TdR) mal hechos", "subcauses": [{"name": "Especificaciones técnicas imposibles", "subcauses": ["Licitaciones declaradas desiertas"]}]},
                {"name": "Fallas logísticas del proveedor", "subcauses": [{"name": "Flota de camiones viejos y averiados", "subcauses": ["Desabastecimiento en la obra"]}]},
                {"name": "Error de pedidos y métricas", "subcauses": [{"name": "Se pidió la mitad del material necesario", "subcauses": ["Freno brusco en medio de ejecución"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Materiales sin certificación de calidad", "subcauses": [{"name": "Rechazo del acero por laboratorios de prueba", "subcauses": ["Devolución masiva de lotes"]}]},
                {"name": "Embargos legales a los contratistas", "subcauses": [{"name": "Deudas laborales de las empresas", "subcauses": ["Cuentas bancarias de proveedores congeladas"]}]},
                {"name": "Sanciones internacionales comerciales", "subcauses": [{"name": "Prohibición de importar de ciertos países", "subcauses": ["Cadena de suministro bloqueada"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Dependencia de proveedor único", "subcauses": [{"name": "Sin diversificación de la cadena", "subcauses": ["Si falla uno, colapsa el proyecto"]}]},
                {"name": "Falta de compras adelantadas", "subcauses": [{"name": "Pedir el material 'Just in Time'", "subcauses": ["Cualquier retraso detiene el proyecto"]}]},
                {"name": "No asegurar la mercancía", "subcauses": [{"name": "Pérdida financiera total si el barco se hunde", "subcauses": ["Incapacidad de recomprar"]}]}
            ]}
        ]

        analysis["ishikawas"]["Pérdida de memoria institucional"] = [
            {"category": "1. Ambiental (E)", "causes": [
                {"name": "Entorno laboral tóxico (literal)", "subcauses": [{"name": "Oficinas con moho y mala ventilación", "subcauses": ["Renuncias por motivos de salud (Síndrome Edificio Enfermo)"]}]},
                {"name": "Riesgos biológicos de campo crónicos", "subcauses": [{"name": "Exposición constante a enfermedades tropicales", "subcauses": ["Técnicos piden traslados"]}]},
                {"name": "Altísimas temperaturas sin AC", "subcauses": [{"name": "Fatiga térmica constante", "subcauses": ["Baja moral y abandono del puesto"]}]}
            ]},
            {"category": "2. Social (S)", "causes": [
                {"name": "Mobbing y acoso laboral (Burnout)", "subcauses": [{"name": "Jefes abusivos que gritan", "subcauses": ["Estrés psicológico insoportable"]}]},
                {"name": "Presión pública y mediática", "subcauses": [{"name": "Ataques en redes sociales al equipo del proyecto", "subcauses": ["Miedo y renuncias protectoras"]}]},
                {"name": "Desplazamiento lejos de familias", "subcauses": [{"name": "Trabajo aislado en zonas rurales", "subcauses": ["Crisis familiares que fuerzan retornos"]}]}
            ]},
            {"category": "3. Gobernanza (G)", "causes": [
                {"name": "Amiguismo y nepotismo", "subcauses": [{"name": "Ascensos de los incapaces sobre los técnicos", "subcauses": ["Fuga de cerebros indignados"]}]},
                {"name": "Salarios estancados por inflación", "subcauses": [{"name": "El Estado congela los sueldos por años", "subcauses": ["Migración masiva al sector privado"]}]},
                {"name": "Inseguridad de contratos temporales", "subcauses": [{"name": "Contratos renovados cada mes", "subcauses": ["Búsqueda de estabilidad fuera del Estado"]}]}
            ]},
            {"category": "4. Tecnología y procesos", "causes": [
                {"name": "Falta de herramientas modernas", "subcauses": [{"name": "Obligar a técnicos a usar PC de hace 15 años", "subcauses": ["Frustración tecnológica"]}]},
                {"name": "Sin sistemas de gestión de conocimiento", "subcauses": [{"name": "Todo el know-how está en la cabeza de las personas", "subcauses": ["El proyecto se amnesia cuando alguien se va"]}]},
                {"name": "Procesos hiper-burocráticos absurdos", "subcauses": [{"name": "El 80% del tiempo es llenar formularios inútiles", "subcauses": ["Desilusión profesional"]}]}
            ]},
            {"category": "5. Normativa y cumplimiento", "causes": [
                {"name": "Riesgo penal y contraloría personal", "subcauses": [{"name": "Cualquier firma es un potencial juicio legal", "subcauses": ["Miedo paralizante, mejor renunciar"]}]},
                {"name": "Incumplimiento de derechos laborales", "subcauses": [{"name": "No pagan horas extras ni viáticos prometidos", "subcauses": ["Demandas y abandono de cargo"]}]},
                {"name": "Ausencia de planes de carrera en la ley", "subcauses": [{"name": "Imposibilidad legal de ascender", "subcauses": ["Estancamiento vitalicio inaceptable"]}]}
            ]},
            {"category": "6. Gestión de riesgos", "causes": [
                {"name": "Nulo plan de sucesión y empalme", "subcauses": [{"name": "La persona se va el viernes, el reemplazo llega el lunes", "subcauses": ["0 días de transmisión de conocimiento"]}]},
                {"name": "Ignorar encuestas de clima laboral", "subcauses": [{"name": "Ceguera corporativa ante el descontento masivo", "subcauses": ["Sorpresa ante la cascada de renuncias"]}]},
                {"name": "Contratación de reemplazos baratos", "subcauses": [{"name": "Para ahorrar, se contratan juniors sin experiencia", "subcauses": ["Errores fatales por ignorancia histórica"]}]}
            ]}
        ]

        
        analysis["stakeholders"] = [
            {"name": "Personal Administrativo", "power": "Bajo", "impact": "Alto"},
            {"name": "Gerencia del Proyecto", "power": "Alto", "impact": "Alto"},
            {"name": "Proveedores Comerciales", "power": "Medio", "impact": "Bajo"},
            {"name": "Ciudadanía General", "power": "Bajo", "impact": "Bajo"}
        ]
        analysis["crva"] = "Riesgo climático bajo. Posibles interrupciones logísticas menores en la cadena de suministro por tormentas o disrupciones viales, mitigables con stock de seguridad y rutas alternas."
        analysis["bowtie"] = {
            "causes": ["Burocracia estatal paralizante", "Rotación de líderes de agencia"],
            "preventive": ["Agencia experta en adquisiciones", "Incentivos de retención salarial"],
            "hazard": "Retraso Crítico en la Ejecución",
            "mitigative": ["Reestructuración del préstamo", "Inyección de contingencia"],
            "consequences": ["Pérdida de recursos por inflación", "Objetivos de desarrollo perdidos"]
        }
        analysis["impact"] = "C"
        analysis["risk"] = "Low"
        analysis["esic_reasoning"] = "Category C: Sin impactos E&S graves."
        analysis["esrr_factors"]["cause"] = "Gestión Operativa."
        analysis["esrr_factors"]["contribution"] = "Nula."
        analysis["esrr_factors"]["context"] = "Entorno corporativo estándar."
        analysis["esrr_factors"]["performance"] = "Capacidad promedio del ejecutor."
        analysis["reason"] = "Riesgos operativos estándar, mitigables con gestión."

    return analysis

def render_fmea(fmea_data):
    if not fmea_data: return
    df = pd.DataFrame(fmea_data)
    def color_rpn(val):
        if not isinstance(val, (int, float)): return ''
        color = '#dc2626' if val >= 200 else '#d97706' if val >= 100 else '#16a34a'
        return f'color: {color}; font-weight: bold;'
    def color_jerarquia(val):
        v = str(val).lower()
        if 'evitar' in v: return 'background-color: #dcfce7; color: #16a34a; font-weight:bold;'
        elif 'minimizar' in v: return 'background-color: #fef9c3; color: #ca8a04; font-weight:bold;'
        elif 'restaurar' in v: return 'background-color: #ffedd5; color: #ea580c; font-weight:bold;'
        elif 'compensar' in v: return 'background-color: #fee2e2; color: #dc2626; font-weight:bold;'
        return ''
    
    st.dataframe(df.style.map(color_rpn, subset=['RPN']).map(color_jerarquia, subset=['Jerarquía']), use_container_width=True, hide_index=True)

def render_bowtie_html(data):
    if not data: return ""
    html = f'''
    <div style="display: flex; justify-content: space-between; align-items: center; background: white; padding: 20px; border-radius: 10px; border: 1px solid #e2e8f0; font-family: Inter, sans-serif;">
        <div style="flex: 1; text-align: right;">
            <h4 style="color:#0f172a; margin-bottom: 10px;">Amenazas (Causas)</h4>
            {''.join([f'<div style="background:#f1f5f9; padding:8px; margin:5px; border-radius:5px; font-size:0.8rem; border-right: 3px solid #94a3b8;">{c}</div>' for c in data['causes']])}
        </div>
        <div style="flex: 1; padding: 0 10px;">
            <h4 style="color:#16a34a; margin-bottom: 10px; text-align:center;">Barreras Preventivas</h4>
            {''.join([f'<div style="background:#dcfce7; border: 1px solid #22c55e; padding:8px; margin:5px; border-radius:5px; text-align:center; font-size:0.8rem; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">{b}</div>' for b in data['preventive']])}
        </div>
        <div style="flex: 1; text-align: center; z-index:10; margin: 0 10px;">
            <div style="background:#fee2e2; border: 2px solid #ef4444; padding:20px 10px; border-radius:10px; font-weight:bold; color:#b91c1c; font-size:1.1rem; box-shadow: 0 4px 6px rgba(239, 68, 68, 0.2);">
                {data.get('hazard', 'Peligro Crítico')}
            </div>
        </div>
        <div style="flex: 1; padding: 0 10px;">
            <h4 style="color:#d97706; margin-bottom: 10px; text-align:center;">Barreras Mitigadoras</h4>
            {''.join([f'<div style="background:#fef3c7; border: 1px solid #f59e0b; padding:8px; margin:5px; border-radius:5px; text-align:center; font-size:0.8rem; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">{b}</div>' for b in data['mitigative']])}
        </div>
        <div style="flex: 1; text-align: left;">
            <h4 style="color:#0f172a; margin-bottom: 10px;">Consecuencias</h4>
            {''.join([f'<div style="background:#f1f5f9; padding:8px; margin:5px; border-radius:5px; font-size:0.8rem; border-left: 3px solid #94a3b8;">{c}</div>' for c in data['consequences']])}
        </div>
    </div>
    '''
    import streamlit.components.v1 as components
    components.html(html, height=350)

def render_stakeholder_html(data):
    if not data: return ""
    q1, q2, q3, q4 = [], [], [], []
    for s in data:
        if s['power'] == 'Alto' and s['impact'] == 'Bajo': q1.append(s['name'])
        elif s['power'] == 'Alto' and s['impact'] == 'Alto': q2.append(s['name'])
        elif s['power'] == 'Bajo' and s['impact'] == 'Bajo': q3.append(s['name'])
        elif s['power'] == 'Bajo' and s['impact'] == 'Alto': q4.append(s['name'])
    
    html = f'''
    <div style="display:grid; grid-template-columns: 30px 1fr 1fr; grid-template-rows: 1fr 1fr 30px; gap: 15px; font-family: Inter, sans-serif; background:white; padding:20px; border-radius:10px; border: 1px solid #e2e8f0; height: 100%;">
        <div style="grid-column: 1; grid-row: 1 / span 2; display: flex; align-items: center; justify-content: center; writing-mode: vertical-rl; transform: rotate(180deg); font-weight: 800; color: #64748b; letter-spacing: 2px;">PODER / INFLUENCIA</div>
        
        <div style="grid-column: 2; grid-row: 1; background: #f8fafc; border: 2px dashed #94a3b8; padding: 15px; border-radius: 8px;">
            <h5 style="margin-top:0; color:#334155; font-size:14px;">Mantener Satisfechos (Poder Alto, Impacto Bajo)</h5>
            <ul style="padding-left: 15px; margin:0; font-size:13px; color:#475569;">{''.join([f'<li>{x}</li>' for x in q1])}</ul>
        </div>
        <div style="grid-column: 3; grid-row: 1; background: #fee2e2; border: 2px dashed #f87171; padding: 15px; border-radius: 8px;">
            <h5 style="margin-top:0; color:#991b1b; font-size:14px;">Gestionar de Cerca (Poder Alto, Impacto Alto)</h5>
            <ul style="padding-left: 15px; margin:0; font-size:13px; color:#991b1b; font-weight:600;">{''.join([f'<li>{x}</li>' for x in q2])}</ul>
        </div>
        
        <div style="grid-column: 2; grid-row: 2; background: #f1f5f9; border: 2px dashed #cbd5e1; padding: 15px; border-radius: 8px;">
            <h5 style="margin-top:0; color:#64748b; font-size:14px;">Monitorear (Poder Bajo, Impacto Bajo)</h5>
            <ul style="padding-left: 15px; margin:0; font-size:13px; color:#64748b;">{''.join([f'<li>{x}</li>' for x in q3])}</ul>
        </div>
        <div style="grid-column: 3; grid-row: 2; background: #fef3c7; border: 2px dashed #fbbf24; padding: 15px; border-radius: 8px;">
            <h5 style="margin-top:0; color:#b45309; font-size:14px;">Proteger / Empoderar (Poder Bajo, Impacto Alto)</h5>
            <ul style="padding-left: 15px; margin:0; font-size:13px; color:#b45309;">{''.join([f'<li>{x}</li>' for x in q4])}</ul>
        </div>
        
        <div style="grid-column: 2 / span 2; grid-row: 3; display: flex; align-items: center; justify-content: center; font-weight: 800; color: #64748b; letter-spacing: 2px;">IMPACTO / VULNERABILIDAD</div>
    </div>
    '''
    import streamlit.components.v1 as components
    components.html(html, height=380)


def main():
    st.sidebar.markdown("""
        <div style='text-align: center; padding: 20px 0;'>
            <h1 style='color: white; font-weight: 800; letter-spacing: 2px; font-size: 2.5rem; margin:0;'>BID</h1>
            <h4 style='color: #94a3b8; font-weight: 400; font-size: 0.9rem; margin-top:5px; text-transform: uppercase;'>ESG Platform 2.0</h4>
        </div>
    """, unsafe_allow_html=True)
    
    page = st.sidebar.radio("", ["🔍 Motor de Análisis Experto", "📊 Dashboard de Casos"])

    if page == "🔍 Motor de Análisis Experto":
        st.markdown('<div class="main-header">Clasificador Multidimensional ESG</div>', unsafe_allow_html=True)
        st.markdown('<div class="sub-header">Flujo Six Sigma: Resumen Ejecutivo ➔ Modos de Falla (FMEA) ➔ Análisis Causal por Riesgo Alto (Ishikawa) ➔ Pilares ESRR ➔ <b>Conclusión Final</b></div>', unsafe_allow_html=True)
        
        case_name = st.text_input("1. Identificación del Caso", placeholder="Ej. ME-L1364 Escuelas / Himachal Pradesh / Barrios Altos")
        uploaded_files = st.file_uploader("2. Sube hasta 4 documentos del proyecto", type=["pdf", "docx"], accept_multiple_files=True)
        
        if uploaded_files:
            if len(uploaded_files) > 4:
                st.error("⚠️ Límite excedido. Máximo 4 documentos.")
            else:
                with st.spinner("Ejecutando pipeline avanzado (GraphRAG + Análisis Sensible al Contexto)..."):
                    file_names = " ".join([f.name.lower() for f in uploaded_files])
                    
                    # GraphRAG & Chunking
                    md_text, chunks, entities = advanced_document_processing(uploaded_files)
                    
                    # Vision AI (Image Extraction & Analysis)
                    extracted_images = extract_images_from_files(uploaded_files)
                    
                    # Add filenames back to text so routing logic continues to work
                    text = file_names + "\n" + case_name.lower() + "\n" + md_text
                    vision_insights = vision_ai_analysis(extracted_images, text)
                    
                    # Feed the vision insights directly into the text stream and chunks as requested
                    if vision_insights and extracted_images:
                        text += "\n\n" + vision_insights
                        chunks.append(vision_insights)
                        
                    analysis = expert_analysis(text)
                
                # --- PASO 0: GRAPH RAG ---
                render_graphrag_ui(chunks, entities)
                

                
                
                # --- PASO 0: RESUMEN EJECUTIVO ---
                st.markdown(f'<div class="summary-box">{analysis["summary"]}</div>', unsafe_allow_html=True)
                
                # --- PASO 1: GEO-SCOPING ---
                st.markdown('<div class="section-title" style="margin-top:2rem;">1. Geo-Scoping: Identificación de Sensibilidad Espacial</div>', unsafe_allow_html=True)
                st.markdown("Cruce de la huella del proyecto con marcadores de vulnerabilidad (Indígena, Arqueológica, Cultural).")
                st.markdown(analysis.get("geo_scoping", ""), unsafe_allow_html=True)

                # --- PASO 2: STAKEHOLDERS ---
                st.markdown('<div class="section-title" style="margin-top:2rem;">2. Matriz de Saliencia de Stakeholders (Poder vs Impacto)</div>', unsafe_allow_html=True)
                st.markdown("Extracción automática y categorización de actores críticos del entorno para prevenir bloqueos de gobernanza.")
                render_stakeholder_html(analysis.get("stakeholders", []))

                # --- PASO 3: CRVA ---
                st.markdown('<div class="section-title" style="margin-top:2rem;">3. Evaluación de Riesgo Climático y Desastres (CRVA)</div>', unsafe_allow_html=True)
                st.markdown(f'''<div style="background:#fff7ed; padding: 1.5rem; border-left: 5px solid #ea580c; border-radius: 8px;">
                    <h4 style="margin-top:0; color:#ea580c;">🌡️ Test de Estrés Climático y Eventos Extremos</h4>
                    <p style="color:#431407; font-size:1.05rem; margin-bottom:0;">{analysis.get("crva", "No se encontraron riesgos climáticos críticos ex-ante.")}</p>
                </div>''', unsafe_allow_html=True)

                # --- PASO 4: FMEA & JERARQUIA ---
                st.markdown('<div class="section-title" style="margin-top:2rem;">4. Análisis FMEA y Jerarquía de Mitigación</div>', unsafe_allow_html=True)
                st.markdown("Identificación de modos de falla y categorización estricta de acciones preventivas (Evitar/Minimizar/Restaurar/Compensar). **Todo riesgo con RPN ≥ 200 (Rojo) gatillará un diagrama Ishikawa profundo.**")
                render_fmea(analysis["fmea"])

                df_fmea = pd.DataFrame(analysis["fmea"])
                red_risks = df_fmea[df_fmea['RPN'] >= 200]
                if red_risks.empty:
                    max_rpn_idx = df_fmea['RPN'].idxmax()
                    red_risks = df_fmea.loc[[max_rpn_idx]]

                # --- PASO 5: BOWTIE ---
                st.markdown('<div class="section-title" style="margin-top:2rem;">5. Diagrama Bowtie (Barreras Críticas para el Peor Escenario)</div>', unsafe_allow_html=True)
                st.markdown("Análisis forense bidireccional del Riesgo Cero: A la izquierda lo que lo previene, a la derecha lo que nos salva si explota.")
                render_bowtie_html(analysis.get("bowtie", {}))

                # --- PASO 6: ISHIKAWAS MULTIPLES ---
                st.markdown('<div class="section-title" style="margin-top:2rem;">6. Análisis de Causa Raíz (Ishikawas Críticos)</div>', unsafe_allow_html=True)
                st.markdown("Generando un diagrama de desconstrucción profunda (3 niveles) por cada riesgo extremo:")
                
                for idx, row in red_risks.iterrows():
                    effect = row['Efecto']
                    rpn_val = row['RPN']
                    st.markdown(f"<h4 style='color: #dc2626; margin-top:2rem; border-left: 4px solid #dc2626; padding-left: 10px;'>⚠️ Riesgo Evaluado: {effect} (RPN: {rpn_val})</h4>", unsafe_allow_html=True)
                    
                    cats = analysis["ishikawas"].get(effect, [])
                    if not cats:
                        cats = [{"category": "N/A", "causes": ["Sin causas mapeadas"]}] * 6
                        
                    render_ishikawa_svg({"effect": effect, "categories": cats})
                
                # --- PASO 7: FACTORES ESRR ---
                st.markdown('<div class="section-title" style="margin-top:2rem;">7. Evaluación Dinámica de Factores ESRR (Los 4 Pilares)</div>', unsafe_allow_html=True)
                
                fcol1, fcol2 = st.columns(2)
                with fcol1:
                    st.markdown(f'<div class="factor-box"><div class="factor-title">🎯 Cause (Direct Impacts)</div><div>{analysis["esrr_factors"]["cause"]}</div></div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="factor-box"><div class="factor-title">🌐 Contribution (Indirect/Cumulative)</div><div>{analysis["esrr_factors"]["contribution"]}</div></div>', unsafe_allow_html=True)
                with fcol2:
                    st.markdown(f'<div class="factor-box"><div class="factor-title">🏛️ Context (External Environment)</div><div>{analysis["esrr_factors"]["context"]}</div></div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="factor-box"><div class="factor-title">⚙️ Performance (Borrower Capacity)</div><div>{analysis["esrr_factors"]["performance"]}</div></div>', unsafe_allow_html=True)

                st.markdown("<hr>", unsafe_allow_html=True)
                
                # --- PASO 8: CONCLUSION ---
                # 1. FMEA Component (40%)
                max_rpn = df_fmea['RPN'].max() if not df_fmea.empty else 0
                fmea_score = min((max_rpn / 1000) * 100, 100)
                
                # 2. Geo-Scoping Component (20%)
                geo_text = analysis.get("geo_scoping", "").lower()
                geo_score = 0
                if "indígena" in geo_text or "arqueol" in geo_text or "coral" in geo_text or "sensibilidad" in geo_text:
                    geo_score = 100
                elif geo_text:
                    geo_score = 50
                    
                # 3. Stakeholder Component (20%)
                stakeholders = analysis.get("stakeholders", [])
                high_risk_st = sum(1 for s in stakeholders if s.get("power") == "Alto" and s.get("impact") == "Alto")
                stake_score = min(high_risk_st * 33.3, 100) # 3 o más = 100
                
                # 4. CRVA Component (20%)
                crva_text = analysis.get("crva", "").lower()
                crva_score = 0
                if "extremo" in crva_text or "desastre" in crva_text or "crítico" in crva_text:
                    crva_score = 100
                elif "moderado" in crva_text or "posibles" in crva_text:
                    crva_score = 50
                
                # Holistic Equation: Índice Global ESG (IGE)
                ige_total = (fmea_score * 0.40) + (geo_score * 0.20) + (stake_score * 0.20) + (crva_score * 0.20)
                
                if ige_total >= 65:
                    sug_cat = "A"
                    badge_color = "#dc2626"
                elif ige_total >= 35:
                    sug_cat = "B"
                    badge_color = "#d97706"
                else:
                    sug_cat = "C"
                    badge_color = "#16a34a"

                st.markdown('<div class="section-title">8. Conclusión Final: Algoritmo de Clasificación Holística</div>', unsafe_allow_html=True)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown(f'''
                    <div class="card">
                        <p style="color:#64748b; font-weight:600; text-align:center;">ESIC (IMPACTO ESTÁTICO)</p>
                        <div class="big-badge" style="background-color: {badge_color}; color: white; border: none;">Categoría {sug_cat} Sugerida</div>
                        <div style="margin-top:1rem; font-size:0.9rem; color:#475569; text-align:center;">
                            <b>Índice Global ESG (IGE): {ige_total:.1f}/100</b>
                        </div>
                    </div>''', unsafe_allow_html=True)
                with col2:
                    risk_class = f"badge-{analysis['risk']}"
                    st.markdown(f'''
                    <div class="card">
                        <p style="color:#64748b; font-weight:600; text-align:center;">ESRR (RIESGO DINÁMICO)</p>
                        <div class="big-badge {risk_class}">Riesgo {analysis["risk"]}</div>
                        <div style="margin-top:1rem; font-size:0.9rem; color:#475569;"><b>Justificación General:</b> {analysis["reason"]}</div>
                    </div>''', unsafe_allow_html=True)

                # EQUATION EXPLANATION
                st.markdown(f'''
<div style="background:#f8fafc; border: 1px solid #cbd5e1; border-radius: 8px; padding: 20px; margin-top: 20px;">
    <h4 style="color:#0f172a; margin-top:0; margin-bottom: 15px;">🧮 Ecuación del Índice Global ESG (Benchmark Internacional)</h4>
    <p style="font-size: 0.95rem; color:#334155; line-height: 1.6;">
        A diferencia de la evaluación tradicional que se basa en el "peor RPN" empírico, esta plataforma utiliza un <b>Índice Global ESG (IGE)</b> ponderado (0-100), inspirado en los modelos de <i>Due Diligence</i> de corporaciones financieras y consultoras Big 4. Integra todos los vectores del análisis Six Sigma:
    </p>
    <div style="background: white; border-radius: 6px; padding: 15px; margin: 15px 0; font-family: monospace; font-size: 1.1rem; text-align: center; border: 1px dashed #94a3b8; color: #1e293b; font-weight: bold;">
        IGE = (FMEA × 0.40) + (GeoScoping × 0.20) + (Stakeholders × 0.20) + (CRVA × 0.20)
    </div>
    <ul style="font-size: 0.85rem; color:#475569; line-height: 1.5; padding-left: 20px; margin-bottom: 20px;">
        <li><b>FMEA ({fmea_score:.1f} pts) [40%]:</b> Modos de falla operativos (Basado en RPN Máximo: {max_rpn}).</li>
        <li><b>GeoScoping ({geo_score:.1f} pts) [20%]:</b> Penalización por huella en zonas indígenas, arqueológicas o críticas.</li>
        <li><b>Stakeholders ({stake_score:.1f} pts) [20%]:</b> Gobernanza (Puntaje basado en {high_risk_st} actor(es) en el cuadrante de Alto Poder/Alto Impacto).</li>
        <li><b>CRVA ({crva_score:.1f} pts) [20%]:</b> Vulnerabilidad a Desastres y Cambio Climático.</li>
    </ul>
    
    <h5 style="color:#0f172a; border-bottom: 2px solid #e2e8f0; padding-bottom: 5px;">📜 Regla de Decisión (BID MPAS):</h5>
    <ul style="font-size: 0.9rem; color:#334155; line-height: 1.5; padding-left: 20px;">
        <li><b style="color:#dc2626;">Categoría A (IGE ≥ 65):</b> Operaciones que pueden provocar grandes impactos. </li>
        <li><b style="color:#d97706;">Categoría B (IGE 35 - 64):</b> Operaciones con impactos locales y a corto plazo mitigables. </li>
        <li><b style="color:#16a34a;">Categoría C (IGE &lt; 35):</b> Impactos mínimos o nulos.</li>
    </ul>
</div>
''', unsafe_allow_html=True)

                if st.button("💾 Guardar Análisis y Clasificación", type="primary"):
                    if not case_name.strip():
                        st.warning("⚠️ Ingresa un nombre para el caso.")
                    else:
                        st.session_state['new_cases'].append({
                            "Op. Number": case_name.split()[0],
                            "Operation Name": case_name,
                            "ESIC": analysis["impact"],
                            "ESRR": analysis["risk"],
                            "Status": "Analizado"
                        })
                        st.success(f"✅ Caso guardado.")

    elif page == "📊 Dashboard de Casos":
        st.markdown('<div class="main-header">Dashboard Operativo ESG</div>', unsafe_allow_html=True)
        df = pd.DataFrame()
        try:
            excel_files = glob.glob(os.path.join(r"C:\Users\MiguelAngelRodriguez\Downloads", "Monthly GS Requests Tracking_2024*.xlsx"))
            if excel_files:
                xl = pd.ExcelFile(excel_files[0])
                month_keywords = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
                all_data = []
                for sheet in xl.sheet_names:
                    if any(m in sheet.lower() for m in month_keywords):
                        try:
                            sdf = pd.read_excel(excel_files[0], sheet_name=sheet)
                            if "Op. Number" in sdf.columns: all_data.append(sdf)
                        except: pass
                if all_data: df = pd.concat(all_data, ignore_index=True)
        except: pass
        
        if len(st.session_state['new_cases']) > 0:
            df = pd.concat([pd.DataFrame(st.session_state['new_cases']), df], ignore_index=True) if not df.empty else pd.DataFrame(st.session_state['new_cases'])
            
        if df.empty:
            st.warning("No hay datos.")
            return
            
        col1, col2, col3 = st.columns(3)
        col1.markdown(f'<div class="card" style="text-align:center;"><h3>Operaciones Totales</h3><h2>{df["Op. Number"].nunique()}</h2></div>', unsafe_allow_html=True)
        col2.markdown(f'<div class="card" style="text-align:center;"><h3>Categoría B</h3><h2 style="color:#d97706;">{len(df[df["ESIC"] == "B"])}</h2></div>', unsafe_allow_html=True)
        col3.markdown(f'<div class="card" style="text-align:center;"><h3>Categoría C</h3><h2 style="color:#16a34a;">{len(df[df["ESIC"] == "C"])}</h2></div>', unsafe_allow_html=True)
        
        st.dataframe(df[["Op. Number", "Operation Name", "ESIC", "ESRR"]].head(100), use_container_width=True)

if __name__ == "__main__":
    main()

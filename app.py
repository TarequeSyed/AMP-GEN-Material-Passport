import streamlit as st
import json
import pandas as pd
from pathlib import Path

st.set_page_config(
    page_title="AMP-GEN Material Passport Dashboard",
    page_icon="🌱",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Title & Description
st.title("🌱 AMP-GEN Material Passport & Carbon Dashboard")
st.markdown("""
Extracts building material inventory and computes EPD-backed carbon estimates from Bill of Quantities (BoQ) scans.
*Supported by the Google Centre for Climate Technology, IIT Roorkee.*
""")

# Load data
data_path = Path("output/passport.json")
meta_path = Path("output/building_meta.json")

if data_path.exists() and meta_path.exists():
    with open(data_path, "r", encoding="utf-8") as f:
        items = json.load(f)
    with open(meta_path, "r", encoding="utf-8") as f:
        meta = json.load(f)
        
    df = pd.DataFrame(items)
    
    # Sidebar
    st.sidebar.header("Building Metadata")
    st.sidebar.markdown(f"**Work:** {meta.get('name_of_work')}")
    st.sidebar.markdown(f"**Organization:** {meta.get('organization')}")
    st.sidebar.markdown(f"**Plinth Area:** {meta.get('plinth_area')}")
    st.sidebar.markdown(f"**Plinth Height:** {meta.get('plinth_height')}")
    st.sidebar.markdown(f"**Depth of Foundation:** {meta.get('depth_of_foundation')}")
    st.sidebar.markdown(f"**Seismic Zone:** {meta.get('seismic_zone')}")
    
    st.sidebar.subheader("Download Outputs")
    # Download buttons
    if Path("output/passport_filled.xlsx").exists():
        with open("output/passport_filled.xlsx", "rb") as f:
            st.sidebar.download_button(
                label="Download Filled Excel Passport",
                data=f,
                file_name="passport_filled.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
    with open("output/passport.json", "rb") as f:
        st.sidebar.download_button(
            label="Download JSON Passport",
            data=f,
            file_name="passport.json",
            mime="application/json"
        )

    # Metrics
    total_carbon = df["embodied_carbon_kg_co2e"].sum()
    total_items = len(df)
    rcc_items = df[df["material_category"] == "Concrete"]
    rcc_carbon = rcc_items["embodied_carbon_kg_co2e"].sum()

    col1, col2, col3 = st.columns(3)
    col1.metric("Total BoQ Items", f"{total_items}")
    col2.metric("Total Embodied Carbon", f"{total_carbon:,.1f} kg CO₂e")
    col3.metric("Concrete Carbon Share", f"{(rcc_carbon/total_carbon)*100:.1f}%" if total_carbon > 0 else "0%")

    # Tabs
    tab1, tab2, tab3 = st.tabs(["📋 Material Passport Sheet", "📊 Carbon Analytics", "🔍 Search & Filter"])
    
    with tab1:
        st.subheader("Material Passport Inventory")
        display_cols = [
            "boq_item_no", "description", "material_category", "material_product",
            "volume_m3", "weight_kg", "density_kg_m3", "gwp_per_kg_co2e", "embodied_carbon_kg_co2e", "comment"
        ]
        # Filter existing columns
        cols_to_show = [c for c in display_cols if c in df.columns]
        st.dataframe(df[cols_to_show], use_container_width=True)
        
    with tab2:
        st.subheader("Material and Carbon Footprint Analytics")
        col_img1, col_img2 = st.columns([2, 1])
        with col_img1:
            st.image("output/visualization.png", use_container_width=True, caption="Material passport & carbon statistics summary.")
        with col_img2:
            st.markdown("""
            ### Carbon Estimation Insights
            * **Mass Conversion**: For volume-based items (`cum`), densities from Indian Standards (IS 456, IS 808, SP 64) were mapped.
            * **Thickness Assumptions**: Area-based items (`sqm`) like plaster and half-brick walls were converted to volume using standard architectural thicknesses.
            * **EPD Source Citations**: CO₂e values are estimated using Indian industrial EPD averages:
              - **Concrete (M15/M20)**: ACC Cement averages
              - **Mild Steel**: Tata Steel India
              - **Aluminium**: Hindalco India
              - **Masonry/Mortar**: UltraTech Cement average
            """)
            
    with tab3:
        st.subheader("Filter Inventory by Material Category")
        selected_category = st.selectbox("Select Material Category", ["All"] + list(df["material_category"].unique()))
        
        filtered_df = df if selected_category == "All" else df[df["material_category"] == selected_category]
        st.dataframe(filtered_df, use_container_width=True)
else:
    st.error("Outputs not found. Please run the pipeline script first using `python src/pipeline.py`.")

import streamlit as st
from weather_tool import generate_weather_report

st.set_page_config(page_title="Weather Report Generator", layout="centered")

st.title("📊 Weather Trend Report Generator")
st.markdown("Generate a 20-year Excel summary of temperature trends by ZIP code.")

with st.form("weather_form"):
    zip_code = st.text_input("Enter ZIP Code", value="10001")
    reference_temp = st.number_input("Reference Temp (°F)", value=85.0)
    submitted = st.form_submit_button("Generate Report")

    if submitted:
        with st.spinner("Crunching weather data..."):
            try:
                output_file = generate_weather_report(zip_code, reference_temp)
                st.success("✅ Report generated successfully!")
                with open(output_file, "rb") as f:
                    st.download_button(
                        label="📥 Download Excel Report",
                        data=f,
                        file_name=output_file,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
            except Exception as e:
                st.error(f"🚨 Error: {str(e)}")
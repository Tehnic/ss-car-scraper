# Marks Dvojeglazovs
# Riga Technical University, 2023/2024

import pandas as pd
import streamlit as st
import openpyxl
import scraper

st.set_page_config(page_title="Car data scraper", page_icon="🚗", layout="wide", initial_sidebar_state="expanded")

st.sidebar.header("Izvēlies režīmu:")
mode = st.sidebar.radio("", ("Real-time", "Test"))

if mode == "Real-time":
    st.sidebar.warning("Uzmanību! Real-time režims var aizņemt dažas minutes vai pat stundas, līdz tiek iegūti visi dati!")
    st.sidebar.warning("Kad dati ir iegūti, dati parādīsies automātiski!")
if mode == "Test":
    st.sidebar.warning("Uzmanību! Test režīms izmanto testa datus no /test-wb/ mapi!")

st.title("Car data scraper")
st.header("Izvēlies marku:")
chosen_mark = st.selectbox("", ("Alfa-Romeo", "Audi", "BMW", "Chevrolet", "Chrysler", "Citroen", "Dacia", "Dodge", "Fiat", "Ford", "Gaz", "Honda", "Hyundai", "Jaguar", "Jeep", "Kia", "Lancia", "Land-Rover", "Lexus", "Mazda", "Mercedes", "Mini", "Mitsubishi", "Nissan", "Opel", "Peugeot", "Porsche", "Renault", "Saab", "Seat", "Skoda", "Smart", "Subaru", "Suzuki", "Toyota", "Uaz", "Vaz", "Volkswagen", "Volvo"))

# under dropdown menu show button "Start" and then send selected mark and mode to generate_file function
if mode == "Real-time":
    if st.button("Start"):
        scraper.generate_file(chosen_mark.lower(), mode)

# after generating file, show button "Show data" and then show data from generated file
if mode == "Test":
    if st.button("Parādīt datus"):
        wb = openpyxl.load_workbook("test-wb/" + chosen_mark.lower() + ".xlsx")
        ws = wb.active
        ws = pd.DataFrame(ws.values)
        st.dataframe(ws)

def show_data(chosen_mark):
    wb = openpyxl.load_workbook("wb/" + chosen_mark.lower() + ".xlsx")
    ws = wb.active
    ws = pd.DataFrame(ws.values)
    st.dataframe(ws)
/* Sustituye office.js dentro del iframe del simulador: usa el Excel simulado de la página padre. */
window.Excel = window.parent.MockExcel.Excel;
window.Office = window.parent.MockExcel.Office;

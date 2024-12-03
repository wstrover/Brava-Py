#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QTableView, QStyledItemDelegate, QCheckBox, QHBoxLayout, QScrollArea, QPushButton, QDialog, QVBoxLayout, QComboBox, QSpinBox, QPushButton, QLabel, QRadioButton, QGroupBox, QFormLayout, QMessageBox
from PyQt5.QtCore import QAbstractTableModel, Qt
from PyQt5.QtGui import QColor
import sys


#df = pd.read_excel('Brava_data.xlsx')
#df = pd.read_excel(r'C:\Users\willi\Desktop\Equity_Univ_1500_Composite_HC (2).xlsx')
#Equity_Univ_1500_Composite_HC (2)

df = pd.read_excel('Equity_Univ_1500_Composite_HC.xlsx')


start_col = 'Act_Shares_%chg_3y'  
start_index = df.columns.get_loc(start_col)


ascending_columns = [
    'Volatil_260d',
    'VALU_MDN_3y',
    'VALU_MDN_90d',
    'Act_Leverage_avg_3y',
    'Act_Shares_%chg_3y'
]


#General % ranking of each numerical value for the company
for col in df.columns[start_index:]:
    if col in ascending_columns:
        df[f'Rank {col}'] = (df[col].rank(pct=True, ascending=False) * 100)
    else:
        df[f'Rank {col}'] = (df[col].rank(pct=True, ascending=True) * 100)



columns_to_show_on_start = [
            'Ticker', 'Name', 'BICS_4_code', 'BICS_1', 'BICS_2', 'BICS_3', 'BICS_4',
            'Fundamental Score', 'Est Mmtm', 'Px Mmtm', 'Value', 'Composite'
        ]
#hiddenColumnsOnStartUp = ['Rank VALU_MDN_90d']

    

    
#Calculates the Fundamental Score
#df['Fundamental Score'] = (df['Rank Act Sales PS %chg 3y']*.2) + (df['Rank Act ebit %chg 3y']*.1) + (df['Rank Act RoIC avg 3y']*.2) + (df['Rank Act FCF mgn avg 3y']*.2) + (df['Rank Act Leverage avg 3y']*.3)
fundamentalWeight = {
    'Rank Act_Shares_%chg_3y' : .1,
    'Rank Act_Sales_%chg_3y' : .05,
    'Rank Act_Sales_PS_%chg_3y' : .05,
    'Rank Act_ebit_%chg_3y' : .1,
    'Rank Act_RoIC_avg_3y' : .2,
    'Rank Act_FCF_mgn_avg_3y' : .2,
    'Rank Act_Leverage_avg_3y' : .3
}
df['Fundamental Score'] = 0
for x in fundamentalWeight:
    df['Fundamental Score'] += (df[x] * fundamentalWeight[x])


# Estimated Momentum
#df['Est Mmtm'] = (df['Rank Est Sales %chg 90d']*.4) + (df['Rank Est ebitda %chg 90d']*.3) + (df['Rank Est EPS %chg 90d']*.3)
estMmtmWeight = {
    'Rank Est_Sales_%chg_90d' : .4,
    'Rank Est_ebitda_%chg_90d' : .3,
    'Rank Est_EPS_%chg_90d' : .3
}
df['Est Mmtm'] = 0
for x in estMmtmWeight:
    df['Est Mmtm'] += (df[x] * estMmtmWeight[x])


#Price Momentum
#df['Px Mmtm'] = (df['Rank Volatil 260d']*.4) + (df['Rank 200 dma % 260d']*.3) + (df['Rank RSI 260']*.3)
pxMomentumWeight = {
    'Rank Volatil_260d': 0.4,
    'Rank 200_dma_%_260d': 0.3,
    'Rank RSI_260': 0.3
}
df['Px Mmtm'] = 0
for x in pxMomentumWeight:
    df['Px Mmtm'] += df[x] * pxMomentumWeight[x]


#Value
#df['Value'] = (df['Rank VALU\n MDN 3y']*.3) + (df['Rank VALU\n MDN 90d']*.7)
valueWeight = {
    'Rank VALU_MDN_3y': 0.3,
    'Rank VALU_MDN_90d': 0.7
}
df['Value'] = 0
for x in valueWeight:
    df['Value'] += df[x] * valueWeight[x]


# Composite
#df['Composite'] = (df['Fundamental Score']*.25) + (df['Est Mmtm']*.35) + (df['Px Mmtm']*.2) + (df['Value']*.2)
compositeWeight = {
    'Fundamental Score': 0.25,
    'Est Mmtm': 0.35,
    'Px Mmtm': 0.2,
    'Value': 0.2
}
df['Composite'] = 0
for x in compositeWeight:
    df['Composite'] += df[x] * compositeWeight[x]

#This function takes in 4 arguements, Example useage. filterFunc(False, 'Composite', 50)
# in this example False meants that it's giving the data in descending order
# 'Composite' is the row it is filtering for
# 50 is the number of companies it is returns

# More advanced Use case example. filterFunc(False, 'Value', 50, filterFunc(False, 'Composite', 50))
# This example uses the fourth arguement which is defaulted to the entire dataset 
# the default is replaced with filterFunc(False, 'Composite', 50) whihc is the filtered top 50 companies by Composite score
# it then filters it descendingly by their 'Value Scores' but because it's returning 50 companies it's still showing the 
# top 50 composite scoring companies

def filterFunc(Asc, Column, Number, Source = df):
    if Asc == True:
        Filtered = Source.nsmallest(Number, Column)
    elif Asc == False:
        Filtered = Source.nlargest(Number, Column)
    return Filtered


#print(df.where(df['BICS_1'] == Technology))
#print(filterFunc(False, 'Composite', 50).where(df['BICS_1'] == 'Technology'))
#print(filterFunc(False, 'Value', 50, filterFunc(False, 'Composite', 50)))

class ColorDelegate(QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super().initStyleOption(option, index)
        try:
            value = float(index.data())
            if value < 0:
                option.palette.setColor(option.palette.Text, QColor('red'))
            else:
                option.palette.setColor(option.palette.Text, QColor('black'))
        except (ValueError, TypeError):
            pass


class DataFrameModel(QAbstractTableModel):
    def __init__(self, data):
        super().__init__()
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._data.columns[section])
            elif orientation == Qt.Vertical:
                return str(self._data.index[section])
        return None


class FilterDialog(QDialog):
    def __init__(self, dataframe, filter_history):
        super().__init__()
        self.setWindowTitle("Filter Options")
        self.setFixedSize(300, 200)
        
        self.dataframe = dataframe
        self.filter_history = filter_history 
        
        self.layout = QVBoxLayout()
        
        self.sort_group = QGroupBox("Sort Order")
        self.sort_layout = QHBoxLayout()
        self.asc_radio = QRadioButton("Ascending")
        self.desc_radio = QRadioButton("Descending")
        self.asc_radio.setChecked(True) 
        self.sort_layout.addWidget(self.asc_radio)
        self.sort_layout.addWidget(self.desc_radio)
        self.sort_group.setLayout(self.sort_layout)
        
        self.column_label = QLabel("Select Column:")
        self.column_combobox = QComboBox()
        self.column_combobox.addItems(dataframe.columns)
        
        self.row_label = QLabel("Number of Rows:")
        self.row_spinbox = QSpinBox()
        self.row_spinbox.setMinimum(1)
        self.row_spinbox.setMaximum(self.dataframe.shape[0])  
        
        self.filter_button = QPushButton("Apply Filter")
        self.filter_button.clicked.connect(self.apply_filter)
        
        self.layout.addWidget(self.sort_group)
        self.layout.addWidget(self.column_label)
        self.layout.addWidget(self.column_combobox)
        self.layout.addWidget(self.row_label)
        self.layout.addWidget(self.row_spinbox)
        self.layout.addWidget(self.filter_button)
        
        self.setLayout(self.layout)

    def apply_filter(self):
        asc = self.asc_radio.isChecked()
        column = self.column_combobox.currentText()
        rows = self.row_spinbox.value()
        
        
        filtered_data = filterFunc(asc, column, rows, self.dataframe)
        
        filter_name = f"Top {rows} companies by {column}" if not asc else f"Lowest {rows} companies by {column}"
        
        
        if self.filter_history:
            history_str = " -- ".join(self.filter_history)
            filter_name = f"{history_str} -- Filtered again - {filter_name}"
        
        self.filter_history.append(filter_name)
        
        
        self.open_filtered_window(filtered_data, filter_name)

    def open_filtered_window(self, dataframe, windowName):
        new_window = MainWindow(dataframe, self.filter_history)
        new_window.setWindowTitle(f"Filtered View - {windowName}")

        new_window.show()
        self.accept()
        new_window.exec_() 
        

class ChangeWeightsDialog(QDialog):
    def __init__(self, dataframe, weight_dicts, table_view):
        super().__init__()
        self.setWindowTitle("Change Weights")
        self.setFixedSize(400, 300)

        self.dataframe = dataframe
        self.weight_dicts = weight_dicts
        self.table_view = table_view  

        self.layout = QVBoxLayout()

        self.column_label = QLabel("Select Composite Column:")
        self.column_combobox = QComboBox()
        self.column_combobox.addItems(weight_dicts.keys())

        self.weights_layout = QVBoxLayout()
        self.weight_inputs = {}

        self.column_combobox.currentTextChanged.connect(self.update_weights)

        self.update_weights()

        self.save_button = QPushButton("Save Changes")
        self.save_button.clicked.connect(self.save_changes)

        self.layout.addWidget(self.column_label)
        self.layout.addWidget(self.column_combobox)
        self.layout.addLayout(self.weights_layout)
        self.layout.addWidget(self.save_button)
        self.setLayout(self.layout)

    def update_weights(self):
        while self.weights_layout.count():
            item = self.weights_layout.takeAt(0)
            if item:
                if item.widget():
                    item.widget().deleteLater() 
                elif item.layout():
                    self.clear_layout(item.layout())

        selected_column = self.column_combobox.currentText()
        if not selected_column:
            return

        current_weights = self.weight_dicts.get(selected_column, {})

        self.weight_inputs.clear()
        for col, weight in current_weights.items():
            row_layout = QHBoxLayout()
            label = QLabel(col)
            spinbox = QSpinBox()
            spinbox.setRange(0, 100)  
            spinbox.setValue(int(weight * 100))  
            row_layout.addWidget(label)
            row_layout.addWidget(spinbox)
            self.weights_layout.addLayout(row_layout)
            self.weight_inputs[col] = spinbox

    def clear_layout(self, layout):
        while layout.count():
            item = layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                self.clear_layout(item.layout())

    def save_changes(self):
        selected_column = self.column_combobox.currentText()
        new_weights = {col: spinbox.value() / 100 for col, spinbox in self.weight_inputs.items()}

        if not abs(sum(new_weights.values()) - 1) < 1e-6:
            QMessageBox.warning(self, "Error", "Weights must sum to 1!")
            return

        self.weight_dicts[selected_column] = new_weights
        self.dataframe[selected_column] = 0

        for col, weight in new_weights.items():
            self.dataframe[selected_column] += self.dataframe[col] * weight

        if self.table_view:
            self.table_view.model().layoutChanged.emit() 

        QMessageBox.information(self, "Success", f"Weights for {selected_column} updated!")
        self.accept()





class IndustrySelectWidget(QWidget):
    
    #  Consumer : Consumer Discretionary, Consumer Staples
    #  Resources : Energy, Materials, Utilities
    #  Health : Health Care
    #  Industrials : Industrials
    #  TNT : Technology, Communications
    #  Financials : Financials, Real Estate
    
    brava_presets = {
        'Consumer': ['Consumer Discretionary', 'Consumer Staples'],
        'Resources': ['Energy', 'Materials', 'Utilities'],
        'Health': ['Health Care'],
        'Industrials': ['Industrials'],
        'TMT': ['Technology', 'Communications'],
        'Financials': ['Financials', 'Real Estate']
    }
    
    def __init__(self, dataframe, update_callback):
        super().__init__()
        self.dataframe = dataframe
        self.update_callback = update_callback

        self.layout = QVBoxLayout()

        self.column_label = QLabel("Select BICS Category:")
        self.column_combobox = QComboBox()
        self.column_combobox.addItems(['No Choice', 'BICS_1', 'BICS_2', 'BICS_3', 'BICS_4', 'Brava Presets'])

        self.value_label_1 = QLabel("Select Industry Value 1:")
        self.value_label_2 = QLabel("Select Industry Value 2:")
        self.value_label_3 = QLabel("Select Industry Value 3:")

        self.value_combobox_1 = QComboBox()
        self.value_combobox_2 = QComboBox()
        self.value_combobox_3 = QComboBox()

        self.apply_button = QPushButton("Apply")
        self.apply_button.clicked.connect(self.apply_filter)

        self.column_combobox.currentTextChanged.connect(self.update_values)

        self.layout.addWidget(self.column_label)
        self.layout.addWidget(self.column_combobox)
        self.layout.addWidget(self.value_label_1)
        self.layout.addWidget(self.value_combobox_1)
        #self.layout.addWidget(self.value_label_2)
        self.layout.addWidget(self.value_combobox_2)
        #self.layout.addWidget(self.value_label_3)s
        self.layout.addWidget(self.value_combobox_3)
        self.layout.addWidget(self.apply_button)

        self.setLayout(self.layout)

        self.update_values()

    def update_values(self):
        selected_column = self.column_combobox.currentText()

        if selected_column == 'Brava Presets':
            self.value_combobox_1.clear()
            self.value_combobox_1.addItem('None')
            self.value_combobox_1.addItems(sorted(self.brava_presets.keys()))

            self.value_combobox_2.clear()
            self.value_combobox_3.clear()

        elif selected_column != 'No Choice':
            cleaned_values = self.dataframe[selected_column].dropna().str.strip()
            unique_values = cleaned_values.unique()

            self.value_combobox_1.clear()
            self.value_combobox_1.addItem('None')
            self.value_combobox_1.addItems(sorted(unique_values.tolist()))

            self.value_combobox_2.clear()
            self.value_combobox_2.addItem('None') 
            self.value_combobox_2.addItems(sorted(unique_values.tolist()))

            self.value_combobox_3.clear()
            self.value_combobox_3.addItem('None') 
            self.value_combobox_3.addItems(sorted(unique_values.tolist()))
        else:
            self.value_combobox_1.clear()
            self.value_combobox_2.clear()
            self.value_combobox_3.clear()

    def apply_filter(self):
        selected_column = self.column_combobox.currentText()
        selected_value_1 = self.value_combobox_1.currentText()
        selected_value_2 = self.value_combobox_2.currentText()
        selected_value_3 = self.value_combobox_3.currentText()

        filters = []

        if selected_column == 'Brava Presets' and selected_value_1 != 'None':
            preset_values = self.brava_presets.get(selected_value_1, [])
            filters.append(self.apply_preset_filter(preset_values))
        else:
            if selected_value_1 != 'None' and selected_value_1 != 'No Choice':
                filters.append(self.dataframe[selected_column] == selected_value_1)
            if selected_value_2 != 'None' and selected_value_2 != 'No Choice':
                filters.append(self.dataframe[selected_column] == selected_value_2)
            if selected_value_3 != 'None' and selected_value_3 != 'No Choice':
                filters.append(self.dataframe[selected_column] == selected_value_3)

        if filters:
            combined_filter = filters[0]
            for filter_condition in filters[1:]:
                combined_filter = combined_filter | filter_condition

            filtered_data = self.dataframe[combined_filter]
            self.update_callback(filtered_data)
        else:
            self.update_callback(self.dataframe)

    def apply_preset_filter(self, preset_values):
        """Helper function to filter based on preset values (OR condition for multiple BICS_1 values)."""
        return self.dataframe['BICS_1'].isin(preset_values)







     

        
class MainWindow(QMainWindow):
    def __init__(self, dataframe, filter_history=None):
        super().__init__()
        self.setWindowTitle("DataFrame Viewer")
        self.resize(1600, 800)

        self.filter_history = filter_history if filter_history else []
        self.dataframe = dataframe

        self.columns_to_show = columns_to_show_on_start
        
        self.weight_dicts = {
            'Fundamental Score': fundamentalWeight,
            'Est Mmtm': estMmtmWeight,
            'Px Mmtm': pxMomentumWeight,
            'Value': valueWeight,
            'Composite': compositeWeight,
        }

        self.table = QTableView()
        self.model = DataFrameModel(dataframe)
        self.table.setModel(self.model)

        delegate = ColorDelegate()
        self.table.setItemDelegate(delegate)

        self.checkbox_layout = QVBoxLayout()
        self.checkboxes = {}
        for col in dataframe.columns:
            checkbox = QCheckBox(col)
            if col in self.columns_to_show:
                checkbox.setChecked(True)
            else:
                checkbox.setChecked(False)
            checkbox.stateChanged.connect(self.update_columns)
            self.checkbox_layout.addWidget(checkbox)
            self.checkboxes[col] = checkbox

        self.update_columns()

        self.industry_select_widget = IndustrySelectWidget(self.dataframe, self.update_table)
        self.checkbox_layout.insertWidget(0, self.industry_select_widget)

        self.scroll_area = QScrollArea()
        scroll_content = QWidget()
        scroll_content.setLayout(self.checkbox_layout)
        self.scroll_area.setWidget(scroll_content)
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setMaximumWidth(200)

        self.toggle_button = QPushButton("Toggle Column Widget")
        self.toggle_button.clicked.connect(self.toggle_column_widget)

        self.filter_button = QPushButton("Filter")
        self.filter_button.clicked.connect(self.open_filter_dialog)

        self.change_weights_button = QPushButton("Change Weights")
        self.change_weights_button.clicked.connect(self.open_change_weights_dialog)

        self.deselect_button = QPushButton("Deselect All Columns")
        self.deselect_button.clicked.connect(self.deselect_all_columns)

        header_layout = QHBoxLayout()
        header_layout.addWidget(self.toggle_button)
        header_layout.addWidget(self.filter_button)
        header_layout.addWidget(self.change_weights_button)

        layout = QVBoxLayout()
        layout.addLayout(header_layout)

        main_layout = QHBoxLayout()
        main_layout.addWidget(self.scroll_area)
        main_layout.addWidget(self.table)

        self.checkbox_layout.addWidget(self.deselect_button)

        layout.addLayout(main_layout)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def update_columns(self):
        selected_columns = [col for col, checkbox in self.checkboxes.items() if checkbox.isChecked()]
        self.model = DataFrameModel(self.dataframe[selected_columns])
        self.table.setModel(self.model)

    def check_and_hide_columns(self):
        
        for column in self.dataframe.columns:
            if column not in self.checkboxes or not self.checkboxes[column].isChecked():
                self.table.setColumnHidden(self.dataframe.columns.get_loc(column), True)
            else:
                self.table.setColumnHidden(self.dataframe.columns.get_loc(column), False)

    
    def update_table(self, filtered_data):
        self.model = DataFrameModel(filtered_data)
        self.table.setModel(self.model)
        self.check_and_hide_columns()
        #self.update_columns()


    def toggle_column_widget(self):
        self.scroll_area.setVisible(not self.scroll_area.isVisible())

    def open_filter_dialog(self):
        filter_dialog = FilterDialog(self.dataframe, self.filter_history)
        filter_dialog.exec_()

    def open_change_weights_dialog(self):
        dialog = ChangeWeightsDialog(self.dataframe, self.weight_dicts, self.table)
        dialog.exec_()
        self.update_columns()
        self.update_table(self.dataframe)

    def deselect_all_columns(self):
        for checkbox in self.checkboxes.values():
            checkbox.setChecked(False)
        self.update_columns()

    def closeEvent(self, event):
        reply = QMessageBox.question(self, 'Quit', 'Are you sure you want to exit?',
                                     QMessageBox.Yes | QMessageBox.No, QMessageBox.Yes)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()



# Revert changed weights to original setting
        
def main():
    app = QApplication(sys.argv)
    window = MainWindow(df.round(5))
    window.show()
    sys.exit(app.exec_())

#if __name__ == "__main__":
main()




# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





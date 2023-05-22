import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

import pandas as pd
import plotly.graph_objects as go
import plotly.io as pio
from PIL import Image, ImageTk
from sklearn.ensemble import GradientBoostingRegressor
from statsmodels.tsa.seasonal import seasonal_decompose


def process_data():
    # Open file dialog to select the Excel file
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])

    if file_path:
        # Read the Excel file
        df = pd.read_excel(file_path, parse_dates=['date'], index_col='date')

        # Perform seasonal decomposition
        decomposition = seasonal_decompose(df['gross_posted'], model='additive', period=12)

        # Prepare the initial data for training
        initial_data = df.loc[:'2022-12-31']
        initial_components = decomposition.loc[:'2022-12-31']

        # Prepare the data for training
        X_train = initial_components[['trend', 'seasonal', 'resid']]
        y_train = initial_data['gross_posted']

        # Create a gradient boosting regressor and fit it to the initial data
        model = GradientBoostingRegressor()
        model.fit(X_train, y_train)

        # Generate the dates for the next 30 days
        next_30_days = pd.date_range(start=df.index[-1] + datetime.timedelta(days=1), periods=30, freq='D')

        realized_values = []
        forecasted_values = []

        for day in next_30_days:
            # Prepare the data for forecasting the current day
            current_components = decomposition.loc[:day]
            X_current_day = current_components[['trend', 'seasonal', 'resid']]

            # Forecast the current day's gross_posted value
            forecast = model.predict(X_current_day)[-1]

            # Update the model with the actual value for the current day
            actual_value = df.loc[day, 'gross_posted']
            X_train.loc[day] = current_components.iloc[-1]
            y_train.loc[day] = actual_value
            model.fit(X_train, y_train)

            # Store the forecasted and actual values
            realized_values.append(actual_value)
            forecasted_values.append(forecast)

        # Create a Plotly figure for the visualization
        fig = go.Figure()
        fig.add_trace(go.Scatter(x=next_30_days, y=realized_values, mode='lines', name='Realized Values'))
        fig.add_trace(go.Scatter(x=next_30_days, y=forecasted_values, mode='lines', name='Forecasted Values'))

        # Save the Plotly figure as an image
        image_path = 'forecast_plot.png'
        pio.write_image(fig, image_path)

        # Display the Plotly figure within the app
        image = Image.open(image_path)
        photo = ImageTk.PhotoImage(image)
        label = tk.Label(app, image=photo)
        label.image = photo
        label.pack()

        # Display projected amounts in a separate window
        proj_window = tk.Toplevel(app)
        proj_window.title('Projected Amounts')
        proj_text = tk.Text(proj_window)
        proj_text.pack()

        # Insert projected amounts in the text box
        proj_text.insert(tk.END, "Projected Amounts for 'gross_posted':\n")
        for day, forecast in zip(next_30_days, forecasted_values):
            proj_text.insert(tk.END, f"{day.strftime('%Y-%m-%d')}: {forecast}\n")

        # Disable text box editing
        proj_text.configure(state='disabled')

        # Copy button to copy projected amounts to clipboard
        copy_button = tk.Button(proj_window, text="Copy", command=lambda: copy_to_clipboard(proj_text))
        copy_button.pack()


def copy_to_clipboard(text_widget):
    selected_text = text_widget.get("1.0", tk.END)
    app.clipboard_clear()
    app.clipboard_append(selected_text)
    messagebox.showinfo("Copied", "Projected amounts copied to clipboard.")


# Create the Tkinter app
app = tk.Tk()
app.title('Excel Data Forecasting')

# Create a button to trigger data processing
process_button = tk.Button(app, text='Process Data', command=process_data)
process_button.pack()

# Run the Tkinter event loop
app.mainloop()

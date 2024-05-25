import tkinter as tk
from tkinter import messagebox, Toplevel, simpledialog
import requests
import openpyxl


def get_available_currencies():
    try:
        response = requests.get('http://api.nbp.pl/api/exchangerates/tables/A/?format=json')
        data = response.json()
        currencies = [rate['code'] for rate in data[0]['rates']]
        return currencies
    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się pobrać dostępnych walut: {e}")
        return []


def get_exchange_rate(currency_code):
    try:
        response = requests.get(f'http://api.nbp.pl/api/exchangerates/rates/A/{currency_code}/?format=json')
        data = response.json()
        return data['rates'][0]['mid']
    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się pobrać kursu walut: {e}")
        return None


def update_exchange_rate_label(*args):
    currency_code = currency_var.get()
    exchange_rate = get_exchange_rate(currency_code)
    if exchange_rate is not None:
        label_exchange_rate.config(text=f"Aktualny kurs {currency_code}/PLN: {exchange_rate:.2f} PLN")


def convert_currency():
    try:
        amount = float(entry_amount.get())
        currency_code = currency_var.get()
        exchange_rate = get_exchange_rate(currency_code)
        if exchange_rate is not None:
            result = amount * exchange_rate
            label_result.config(text=f"{amount} {currency_code} = {result:.2f} PLN")
    except ValueError:
        messagebox.showerror("Błąd", "Proszę wprowadzić prawidłową kwotę.")


def show_currency_table():
    top = Toplevel(root)
    top.title("Kursy walut")

    all_currencies = set(fixed_currencies + additional_currencies)
    rows = []

    for currency in all_currencies:
        exchange_rate = get_exchange_rate(currency)
        if exchange_rate is not None:
            row = f"{currency}: {exchange_rate:.2f} PLN"
        else:
            row = f"{currency}: Błąd pobierania kursu"
        rows.append(row)

    for i, row in enumerate(rows):
        label = tk.Label(top, text=row, font=("Helvetica", 12))
        label.grid(row=i, column=0, padx=10, pady=5)


def add_currency():
    new_currency = simpledialog.askstring("Dodaj walutę", "Podaj kod waluty:")
    if new_currency:
        new_currency = new_currency.upper()
        if new_currency in available_currencies:
            if new_currency not in additional_currencies:
                additional_currencies.append(new_currency)
                messagebox.showinfo("Sukces", f"Waluta {new_currency} została dodana.")
            else:
                messagebox.showinfo("Informacja", f"Waluta {new_currency} już została dodana.")
        else:
            messagebox.showerror("Błąd", f"Waluta {new_currency} nie jest dostępna w NBP.")


def save_to_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Kursy Walut"
    ws.append(["Waluta", "Kurs (PLN)"])

    all_currencies = set(fixed_currencies + additional_currencies)

    for currency in all_currencies:
        exchange_rate = get_exchange_rate(currency)
        if exchange_rate is not None:
            ws.append([currency, exchange_rate])
        else:
            ws.append([currency, "Błąd pobierania kursu"])

    try:
        wb.save("kursy_walut.xlsx")
        messagebox.showinfo("Sukces", "Kursy walut zostały zapisane do pliku kursy_walut.xlsx")
    except Exception as e:
        messagebox.showerror("Błąd", f"Nie udało się zapisać pliku: {e}")


def show_custom_conversion():
    top = Toplevel(root)
    top.title("Przelicznik walut")

    def update_conversion_rate_label(*args):
        from_currency = from_currency_var.get()
        to_currency = to_currency_var.get()

        if from_currency == 'PLN':
            exchange_rate = get_exchange_rate(to_currency)
            if exchange_rate is not None:
                label_conversion_rate.config(text=f"Aktualny kurs {to_currency}/PLN: {exchange_rate:.2f} PLN")
            else:
                label_conversion_rate.config(text=f"Błąd pobierania kursu {to_currency}")
        elif to_currency == 'PLN':
            exchange_rate = get_exchange_rate(from_currency)
            if exchange_rate is not None:
                label_conversion_rate.config(text=f"Aktualny kurs {from_currency}/PLN: {exchange_rate:.2f} PLN")
            else:
                label_conversion_rate.config(text=f"Błąd pobierania kursu {from_currency}")
        else:
            exchange_rate_from = get_exchange_rate(from_currency)
            exchange_rate_to = get_exchange_rate(to_currency)
            if exchange_rate_from is not None and exchange_rate_to is not None:
                label_conversion_rate.config(
                    text=f"Aktualny kurs {from_currency}/{to_currency}: {exchange_rate_from / exchange_rate_to:.4f} {to_currency}")
            else:
                label_conversion_rate.config(text=f"Błąd pobierania kursów {from_currency} lub {to_currency}")

    label_from_currency = tk.Label(top, text="Waluta źródłowa:")
    label_from_currency.grid(row=1, column=0, padx=10, pady=10)
    from_currency_var = tk.StringVar(value='EUR')
    from_currency_var.trace('w', update_conversion_rate_label)
    dropdown_from_currency = tk.OptionMenu(top, from_currency_var, *available_currencies)
    dropdown_from_currency.grid(row=1, column=1, padx=10, pady=10)

    label_to_currency = tk.Label(top, text="Waluta docelowa:")
    label_to_currency.grid(row=2, column=0, padx=10, pady=10)
    to_currency_var = tk.StringVar(value='PLN')
    to_currency_var.trace('w', update_conversion_rate_label)
    dropdown_to_currency = tk.OptionMenu(top, to_currency_var, *available_currencies)
    dropdown_to_currency.grid(row=2, column=1, padx=10, pady=10)

    label_amount = tk.Label(top, text="Kwota:")
    label_amount.grid(row=3, column=0, padx=10, pady=10)
    entry_amount = tk.Entry(top)
    entry_amount.grid(row=3, column=1, padx=10, pady=10)

    label_conversion_rate = tk.Label(top, text="", font=("Helvetica", 12))
    label_conversion_rate.grid(row=0, column=0, columnspan=2, padx=10, pady=10)

    label_result = tk.Label(top, text="", font=("Helvetica", 12))
    label_result.grid(row=5, column=0, columnspan=2, pady=10)

    def convert_custom_currency():
        try:
            amount = float(entry_amount.get())
            from_currency = from_currency_var.get()
            to_currency = to_currency_var.get()

            if from_currency == 'PLN':
                exchange_rate = get_exchange_rate(to_currency)
                if exchange_rate is not None:
                    result = amount / exchange_rate
                    label_result.config(text=f"{amount} PLN = {result:.2f} {to_currency}")
            elif to_currency == 'PLN':
                exchange_rate = get_exchange_rate(from_currency)
                if exchange_rate is not None:
                    result = amount * exchange_rate
                    label_result.config(text=f"{amount} {from_currency} = {result:.2f} PLN")
            else:
                exchange_rate_from = get_exchange_rate(from_currency)
                exchange_rate_to = get_exchange_rate(to_currency)
                if exchange_rate_from is not None and exchange_rate_to is not None:
                    result = amount * exchange_rate_from / exchange_rate_to
                    label_result.config(text=f"{amount} {from_currency} = {result:.2f} {to_currency}")
        except ValueError:
            messagebox.showerror("Błąd", "Proszę wprowadzić prawidłową kwotę.")

    button_convert = tk.Button(top, text="Przelicz", command=convert_custom_currency)
    button_convert.grid(row=4, column=0, columnspan=2, pady=10)

    update_conversion_rate_label()


root = tk.Tk()
root.title("Przelicznik walut na PLN")

available_currencies = get_available_currencies()
if not available_currencies:
    available_currencies = ['EUR']

fixed_currencies = ['EUR', 'USD', 'GBP', 'CHF', 'JPY']
additional_currencies = []


label_exchange_rate = tk.Label(root, text="Pobieranie kursu...", font=("Helvetica", 14))
label_exchange_rate.grid(row=0, column=0, columnspan=3, padx=10, pady=10)

currency_var = tk.StringVar(value='EUR')
currency_var.trace('w', update_exchange_rate_label)

label_currency = tk.Label(root, text="Wybierz walutę:")
label_currency.grid(row=1, column=0, padx=10, pady=10)

dropdown_currency = tk.OptionMenu(root, currency_var, *available_currencies)
dropdown_currency.grid(row=1, column=1, padx=10, pady=10)

label_amount = tk.Label(root, text="Kwota:")
label_amount.grid(row=2, column=0, padx=10, pady=10)

entry_amount = tk.Entry(root)
entry_amount.grid(row=2, column=1, padx=10, pady=10)

button_convert = tk.Button(root, text="Przelicz", command=convert_currency)
button_convert.grid(row=3, column=0, columnspan=2, pady=10)

button_show_table = tk.Button(root, text="Pokaż kursy walut", command=show_currency_table)
button_show_table.grid(row=4, column=0, columnspan=2, pady=10)

button_add_currency = tk.Button(root, text="Dodaj walutę", command=add_currency)
button_add_currency.grid(row=5, column=0, columnspan=2, pady=10)

button_save_to_excel = tk.Button(root, text="Zapisz do Excel", command=save_to_excel)
button_save_to_excel.grid(row=6, column=0, columnspan=2, pady=10)

button_custom_conversion = tk.Button(root, text="Przelicz waluty", command=show_custom_conversion)
button_custom_conversion.grid(row=7, column=0, columnspan=2, pady=10)

label_result = tk.Label(root, text="", font=("Helvetica", 12))
label_result.grid(row=8, column=0, columnspan=2, pady=10)

update_exchange_rate_label()


root.mainloop()

import pandas as pd
import os

# Verificar se o arquivo já existe
if os.path.exists('precos.xlsx'):
    print("Arquivo precos.xlsx já existe. Não será substituído.")
else:
    # Dados dos preços
    pizzas_data = {
        "Produto": ["Muçarela", "Calabresa", "Presunto", "Mista", "Bacon",
                    "Marguerita", "Portuguesa", "Frango", "Calabacon",
                    "Calabresa Apimentada", "Quatro Queijos", "Frango com Requeijao",
                    "Nordestina", "Frango Cremoso", "Siciliana", "Peito de Peru",
                    "À Moda da Casa", "Paraense", "Filé", "Camarão Rosa",
                    "Filé com Bacon", "Savino Especial", "Brigadeiro", "Disqueti"],
        "Preço": [45, 45, 45, 50, 50, 50, 50, 50, 50, 50, 50, 55, 60, 60, 65, 65, 65, 65, 65, 70, 70, 70, 35, 35]
    }
    
    bordas_data = {
        "Produto": ["Catupiry", "Chocolate", "Cheddar", "Requeijão"],
        "Preço": [16, 11, 14, 12]
    }
    
    adicionais_data = {
        "Produto": ["Ovo", "Cebola", "Tomate", "Pimentão Verde", "Jambu", "Bacon",
                    "Calabresa", "Muçarela", "Frango", "Requeijão", "Cheddar",
                    "Peito de Peru", "Catupiry"],
        "Preço": [2, 2, 3, 3, 7, 8, 8, 10, 10, 10, 10, 11, 12]
    }
    
    refrigerantes_data = {
        "Produto": ["Coca-cola 1L", "Tuchaua 1L"],
        "Preço": [9.0, 8.0]
    }
    
    # Criar arquivo Excel com múltiplas abas
    with pd.ExcelWriter('precos.xlsx', engine='openpyxl') as writer:
        pd.DataFrame(pizzas_data).to_excel(writer, sheet_name='Pizzas', index=False)
        pd.DataFrame(bordas_data).to_excel(writer, sheet_name='Bordas', index=False)
        pd.DataFrame(adicionais_data).to_excel(writer, sheet_name='Adicionais', index=False)
        pd.DataFrame(refrigerantes_data).to_excel(writer, sheet_name='Refrigerantes', index=False)
    
    print("Arquivo precos.xlsx criado com sucesso!")
    print("Abas criadas: Pizzas, Bordas, Adicionais, Refrigerantes")

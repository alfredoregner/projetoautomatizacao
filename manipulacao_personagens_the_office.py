# Lista de personagens (em formato Python - semelhante ao JSON)
personagens = [
    {
        "id": 1,
        "nome": "Michael Scott",
        "cargo": "Gerente Regional",
        "departamento": "Administração",
        "idade": 45,
        "data_nascimento": "1965-03-15",
        "email": "michael.scott@dundermifflin.com",
        "ativo": True
    },
    {
        "id": 2,
        "nome": "Dwight Schrute",
        "cargo": "Assistente do Gerente Regional",
        "departamento": "Vendas",
        "idade": 38,
        "data_nascimento": "1977-01-20",
        "email": "dwight.schrute@dundermifflin.com",
        "ativo": True
    },
    # ... outros personagens ...
]

# 1️⃣ Acessar um personagem (por índice)
print("Primeiro personagem:")
print(personagens[0])

# 2️⃣ Adicionar novo personagem
novo_personagem = {
    "id": 16,
    "nome": "Toby Flenderson",
    "cargo": "RH",
    "departamento": "RH",
    "idade": 42,
    "data_nascimento": "1979-02-10",
    "email": "toby.flenderson@dundermifflin.com",
    "ativo": True
}
personagens.append(novo_personagem)

# 3️⃣ Alterar cargo de Pam Beesly
for p in personagens:
    if p["nome"] == "Pam Beesly":
        p["cargo"] = "Designer Gráfica"

# 4️⃣ Excluir personagem (ex: Ryan Howard)
personagens = [p for p in personagens if p["nome"] != "Ryan Howard"]

# Mostrar lista final
print("\nLista atualizada de personagens:")
for p in personagens:
    print(p["nome"], "-", p["cargo"])

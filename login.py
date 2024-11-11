import streamlit as st

# Dicionário com usuários e senhas
usuarios = {"usuario@itau": "senha1", "usuario2@bv": "senha2"}

# Função para autenticar usuário
def autenticar(usuario, senha):
    return usuarios.get(usuario) == senha

# Função para carregar o dashboard dependendo do domínio do usuário
def login():
    st.logo("https://finchsolucoes.com.br/img/fefdd9df-1bd3-4107-ab22-f06d392c1f55.png")
    st.sidebar.header("Login")
    usuario = st.sidebar.text_input("Usuário")
    senha = st.sidebar.text_input("Senha", type="password")

    if st.sidebar.button("Entrar"):
        if autenticar(usuario, senha):
            st.session_state.logado = True
            st.session_state.usuario_logado = usuario  # Armazena o usuário logado
            st.sidebar.success("Login bem-sucedido!")
            return True  # Retorna True se o login for bem-sucedido
        else:
            st.sidebar.error("Usuário ou senha incorretos.")
    
    return False  # Retorna False se o login falhar

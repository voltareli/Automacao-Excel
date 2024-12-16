import React, { useState } from "react";
import axios from "axios";
import './App.scss';

function App() {

    

  

  

  const [arquivo, setArquivo] = useState(null);
  const [mensagem, setMensagem] = useState("");
  const [carregando, setCarregando] = useState(false);

  const handleFileChange = (e) => {
    setArquivo(e.target.files[0]);
  };

  const handleUpload = async (e) => {
    e.preventDefault();

    
    if (!arquivo) {
      setMensagem("Por favor, selecione um arquivo Excel.");
      return;
    }

    const formData = new FormData();
    formData.append("arquivo", arquivo);

    try {
      setCarregando(true);
      setMensagem("Enviando arquivo...");

      const resposta = await axios.post("http://localhost:5030/enviar-arquivo", formData, {
        headers: {
          "Content-Type": "multipart/form-data",
        },
      });

      setMensagem(resposta.data);
    } catch (erro) {
      setMensagem("Erro ao enviar o arquivo: " + erro.message);
    } finally {
      setCarregando(false);
    }
  };
  return (
    <div className="App">
        
      <h1>Automação do Ministério da fazenda</h1>
      <form onSubmit={handleUpload}>
        <div>
          <label htmlFor="arquivo">Selecione o arquivo Excel:</label>
          <input type="file" id="arquivo" accept=".xlsx" onChange={handleFileChange} />
        </div>
        <button type="submit" disabled={carregando}>
          {carregando ? "Processando..." : "Enviar Arquivo"}
        </button>
      </form>
      {mensagem && <p>{mensagem}</p>}

    </div>
  );
}


export default App;

const puppeteer = require("puppeteer");
const express = require("express");
const multer = require("multer");
const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");





const app = express();
const upload = multer({ dest: "uploads" });
const PORT = 5030;

const lerPrimeiraColuna = (caminho) => {
  const planilha = xlsx.readFile(caminho);
  const aba = planilha.SheetNames[0];
  const dados = xlsx.utils.sheet_to_json(planilha.Sheets[aba], { header: 1 });

  const primeiraColuna = dados.map((linha) => linha[0]);

  return primeiraColuna;
};

const iniciarNavegador = async () => {
  const navegador = await puppeteer.launch({
    headless: false,
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });

  const pagina = await navegador.newPage();

  await pagina.setViewport({
    width: 1024,
    height: 768,
  });

  return { navegador, pagina };
};

const aguardarURLCorreta = async (pagina, urlEsperada) => {
  console.log(`Aguardando a navegação manual para a URL: ${urlEsperada}`);
  await pagina.waitForFunction(
    (url) => window.location.href === url,
    { timeout: 0 },
    urlEsperada
  );
  console.log("Navegação para a URL esperada detectada!");
};

const executarAutomacao = async (codigoNota, pagina) => {
  try {
    if (!codigoNota || typeof codigoNota !== "string") {
      throw new Error("O código da nota não é válido.");
    }

    await pagina.waitForSelector('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]', {
      visible: true,
      timeout: 10000,
    });

  
    await pagina.evaluate((codigo) => {
      const input = document.querySelector('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]');
      if (input) {
        input.value = codigo;
        input.dispatchEvent(new Event('input', { bubbles: true }));
      }
    }, codigoNota);

    await pagina.waitForSelector('[value="Salvar Nota"]', { visible: true, timeout: 1000});
    await pagina.click('[value="Salvar Nota"]');

    await pagina.waitForNavigation({ waitUntil: "networkidle2", timeout: 1000 });

    await pagina.click('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]');
    await pagina.evaluate(() => {
      const input = document.querySelector('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]');
      if (input) {
        input.value = "";
        input.dispatchEvent(new Event('input', { bubbles: true }));
      }
    });
    await pagina.click('[title="Digite ou Utilize um leitor de código de barras ou QRCode"]');

  await new Promise(resolve => setTimeout(resolve, 2000)); 

} catch (erro) {
  console.error(`Erro no processo: ${erro}`);
}
};

const excluirArquivosAntigos = (diretorio, tempoLimite) => {
  const diretorioUploads = path.resolve(__dirname, diretorio);

  fs.readdirSync(diretorioUploads).forEach((file) => {
    const caminhoArquivo = path.join(diretorioUploads, file);
    const stats = fs.statSync(caminhoArquivo);
    const tempoCriacao = stats.birthtimeMs;
    const tempoAtual = Date.now();

    if (tempoAtual - tempoCriacao > tempoLimite) {
      fs.unlinkSync(caminhoArquivo);
      console.log(`Arquivo deletado (mais de 24 horas): ${caminhoArquivo}`);
    }
  });
};

app.post("/enviar-arquivo", upload.single("arquivo"), async (req, res) => {
  const arquivo = req.file;

  if (!arquivo) {
    return res.status(400).send("Nenhum arquivo enviado.");
  }

  try {
    const primeiraColuna = lerPrimeiraColuna(arquivo.path);

    const { navegador, pagina } = await iniciarNavegador();

   
    const urlInicial = "https://www.nfp.fazenda.sp.gov.br/login.aspx?ReturnUrl=%2fEntidadesFilantropicas%2fCadastroNotaEntidade.aspx";
    await pagina.goto(urlInicial, { waitUntil: "domcontentloaded" });


    const urlEsperada = "https://www.nfp.fazenda.sp.gov.br/EntidadesFilantropicas/ListagemNotaEntidade.aspx";
    await aguardarURLCorreta(pagina, urlEsperada);


    for (const codigoNota of primeiraColuna) {
     
      const codigoNotaStr = String(codigoNota);


      if (!codigoNotaStr || codigoNotaStr.length !== 44 || isNaN(codigoNotaStr)) {
        console.log(`Valor inválido para Código da Nota: ${codigoNotaStr}`);
        continue;
      }

      await executarAutomacao(codigoNotaStr, pagina);
    }

    const tempoLimite = 24 * 60 * 60 * 1000;
    excluirArquivosAntigos("./uploads", tempoLimite);

    await navegador.close();



    res.send("Automação concluída com sucesso!");
  } catch (erro) {
    console.error("Erro no processo:", erro);
    res.status(500).send("Erro na automação.");
  }
});

app.listen(PORT, () => {
  console.log(`Servidor rodando na porta ${PORT}`);
});
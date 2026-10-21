import java.io.File;
import java.io.FileInputStream;
import java.nio.file.FileSystems;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardWatchEventKinds;
import java.nio.file.WatchEvent;
import java.nio.file.WatchKey;
import java.nio.file.WatchService;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class Main{
    private static String caminho_pasta = "\\\\192.168.1.10\\DatabaseFolder\\01_XML\\IMPORTADOS";
    private static String caminho_arquivo = "RELACAO-MODELOS.xlsx";
    private static String caminho_saida = "\\\\192.168.1.10\\DatabaseFolder\\relatorio_atualizado_1.xlsx";
    
    public static void main(String[] args) throws Exception {

        Map<String, String> codigosExcel = codigosExcel(caminho_arquivo); // pega os codigos do excel na primeira função que chama

        Path diretorioPath = Paths.get(caminho_pasta); // cria o path com a pasta

        File pasta = new File(caminho_pasta); // cria o arquivo com a pasta
        if(!pasta.exists() || !pasta.isDirectory()){
            System.err.println("A pasta não existe ou não é um diretório.");
            return;
        }


        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance(); // cria o builder 
        DocumentBuilder builder = factory.newDocumentBuilder(); // cria o builder

        
        // Carrega os pedidos que já estão no Excel para não duplicar
        Set<String> pedidosExistentes = ExportadorExcel.pedidosExistentes(caminho_saida);
        System.out.println("Pedidos já registrados no relatório: " + pedidosExistentes.size());

        // Processa os arquivos XML que já existem na pasta antes de iniciar o monitoramento
        File[] arquivosExistentes = pasta.listFiles((dir, nome) -> nome.toLowerCase().endsWith(".xml"));
        if(arquivosExistentes != null && arquivosExistentes.length > 0){
            System.out.println("Verificando " + arquivosExistentes.length + " arquivos XML existentes na pasta...");
            Map<String, Map<String, ExportadorExcel.DadosGrupo>> dadosIniciais = new HashMap<>();

            for(File arq : arquivosExistentes){
                String nomeArquivo = arq.getName();
                String pedido = nomeArquivo.length() >= 5 ? nomeArquivo.substring(0, 5) : nomeArquivo;

                if(pedidosExistentes.contains(pedido)){
                    continue;
                }

                Map<String, ExportadorExcel.DadosGrupo> contagem = compararCodigos(arq, builder, codigosExcel);
                if(!contagem.isEmpty()){
                    dadosIniciais.put(pedido, contagem);
                }
            }

            if(!dadosIniciais.isEmpty()){
                System.out.println("Exportando " + dadosIniciais.size() + " pedidos encontrados para o Excel...");
                boolean sucesso = ExportadorExcel.exportar(dadosIniciais, caminho_saida);
                if(sucesso){
                    pedidosExistentes.addAll(dadosIniciais.keySet());
                    System.out.println("Relatorio inicial gerado/atualizado com sucesso!\n");
                } else {
                    System.err.println("Atenção: Não foi possível salvar no Excel (arquivo aberto ou bloqueado). Os pedidos serão reprocessados na próxima execução.\n");
                }
            } else {
                System.out.println("Todos os arquivos existentes já estavam no relatório.\n");
            }
        }

        // 1. Cria e registra o WatchService no diretório de rede
        WatchService watchService = FileSystems.getDefault().newWatchService();
        diretorioPath.register(watchService, StandardWatchEventKinds.ENTRY_CREATE);

        System.out.println(">> Monitoramento ATIVO na pasta de rede! Aguardando novos arquivos XML... (Ctrl+C para encerrar)\n");

        while(true){
            WatchKey key = watchService.take(); // faz com que o programa fique dormindo até entrar um novo arquivo
            Map<String, Map<String, ExportadorExcel.DadosGrupo>> dadosNovos = new HashMap<>(); 

            /* serve para ignorar eventos de erro/overflow*/
            for(WatchEvent<?> event : key.pollEvents()){
                WatchEvent.Kind<?> kind = event.kind();
                
                if(kind == StandardWatchEventKinds.OVERFLOW){
                    continue;
                }
            

                // Pega o nome do arquivo recém-criado
                Path nomeRelativo = (Path) event.context();
                String nomeArquivo = nomeRelativo.toString();
                
                if(nomeArquivo.toLowerCase().endsWith(".xml")){
                    String pedido = nomeArquivo.length() >= 5 ? nomeArquivo.substring(0, 5) : nomeArquivo;
                    
                    if(pedidosExistentes.contains(pedido)){
                        System.out.println("Arquivo ignorado (pedido já no relatório): " + nomeArquivo);
                        continue;
                    }

                    // Aguarda 1 segundo para garantir que o arquivo terminou de ser copiado pela rede
                    Thread.sleep(1000);
                    File arquivoCompleto = diretorioPath.resolve(nomeRelativo).toFile();

                    if(arquivoCompleto.exists() && arquivoCompleto.isFile()){
                        System.out.println("Processando novo arquivo detectado: " + nomeArquivo);
                        Map<String, ExportadorExcel.DadosGrupo> contagem = compararCodigos(arquivoCompleto, builder, codigosExcel);
                        
                        if(!contagem.isEmpty()){
                            dadosNovos.put(pedido, contagem);
                        }
                    }
                }
            }

            /* Salva no Excel apenas o(s) novo(s) arquivo(s) */
            if(!dadosNovos.isEmpty()){
                boolean sucesso = ExportadorExcel.exportar(dadosNovos, caminho_saida);
                if(sucesso){
                    pedidosExistentes.addAll(dadosNovos.keySet());
                    System.out.println("Excel atualizado com sucesso!\n");
                } else {
                    System.err.println("Atenção: Não foi possível atualizar o Excel no momento (arquivo aberto por outro processo).\n");
                }
            }

            /* Reseta a chave para ficar pronto para o próximo evento*/
            boolean valido = key.reset();

            if(!valido){
                System.err.println("O diretório monitorado não está mais acessível.");
                break;
            }
        }
    }

    // Função que compara os códigos do xml com os do excel e acumula quantidade e valor total
    private static Map<String, ExportadorExcel.DadosGrupo> compararCodigos(File xml, DocumentBuilder builder, Map<String, String> codigosExcel){

        Map<String, ExportadorExcel.DadosGrupo> dadosPorGrupo = new HashMap<>();

        try{
            Document doc = builder.parse(xml); // cria o documento e parseia o xml
            doc.getDocumentElement().normalize(); // serve para organizar o xml

            NodeList listaItens = doc.getElementsByTagName("ITEM"); // pega todos os itens da 'tag' item

            for(int i = 0; i< listaItens.getLength(); i++){
                Node node = listaItens.item(i); // pega o item de cada laço

                if(node.getNodeType() == Node.ELEMENT_NODE){ // verifica se é um elemento, importante para não pegar comentarios ou outros elementos
                    Element elemento = (Element) node; // transforma para Element, obriga o node ser um elemento

                    String codigoXML =  elemento.getAttribute("REFERENCIA").trim().toUpperCase(); // cria uma string chamada referencia e pega o atributo REFERENCIA do item

                    String grupo = null;
                    if(codigosExcel.containsKey(codigoXML)){ // se o codigoXML existir no excel
                        grupo = codigosExcel.get(codigoXML);
                    } else if(codigoXML.startsWith("ES")){ // especiais que não estão no excel (ES, ES01, ES02, ES08, etc.)
                        grupo = "ESPECIAIS";
                    }

                    if(grupo != null){
                        double precoItem = 0.0;
                        String precoStr = elemento.getAttribute("PRECO_TOTAL").trim().replace(",", ".");
                        if(!precoStr.isEmpty()){
                            try {
                                precoItem = Double.parseDouble(precoStr);
                            } catch(NumberFormatException e){
                                precoItem = 0.0;
                            }
                        }

                        dadosPorGrupo.computeIfAbsent(grupo, k -> new ExportadorExcel.DadosGrupo()).adicionar(1, precoItem);
                    }
                }
            }

            // percorre a lista de grupos e imprime a contagem e valor total
            for(Map.Entry<String, ExportadorExcel.DadosGrupo> entrada : dadosPorGrupo.entrySet()){
                System.out.printf("O grupo %s tem %d item(ns) - Total: R$ %.2f%n", 
                    entrada.getKey(), entrada.getValue().getQuantidade(), entrada.getValue().getValorTotal());
            }
            
        }

        catch(Exception e){
            System.err.println("Erro ao processar o XML " + xml.getName() + ": " + e.getMessage());
        }

        return dadosPorGrupo;
    }


    // Lê a coluna "Item" do Excel e armazena no Set
    private static Map<String, String> codigosExcel(String caminho){ 

        Map<String, String> codigos = new HashMap<>(); // Cria a lista de codigos do excel

        try(FileInputStream file = new FileInputStream(new File(caminho)); Workbook workbook = WorkbookFactory.create(file)){
            Sheet aba = workbook.getSheetAt(0); // pega a primeira aba do excel
            DataFormatter formatter = new DataFormatter();

            for(Row linha : aba){ // para cada linha na aba
                if (linha.getRowNum() == 0) continue; // pula o cabeçalho

                Cell celulaCodigo = linha.getCell(0); // pega a coluna 0, ou seja, a primeira coluna
                Cell celulaGrupo = linha.getCell(2);  // pega a coluna 2, ou seja, a terceira coluna
                
                if(celulaCodigo != null){ // verifica se a celula não está vazia
                    String codigo = formatter.formatCellValue(celulaCodigo).trim(); // pega o valor da celula e remove espaços extras
                    String grupo = formatter.formatCellValue(celulaGrupo).trim();
                    
                    if(!codigo.isEmpty()) {
                        codigos.put(codigo.toUpperCase(), grupo); // Chave = Código, Valor = Grupo
                    }
                }
            }
        }
        catch(Exception e){
            System.err.println("Erro ao ler o arquivo Excel: " + e.getMessage());
            e.printStackTrace();
            //return null; // retorna null se houver erro
        }
        
        return codigos; // retorna a lista de codigos
    }
}








/* 
    public static void main(String[] args) throws Exception {

        // cria o File com o nome da pasta
        File pasta = new File(caminho_pasta);
        
        // verifica se a pasta existe e se é um diretório
        if(pasta.exists() && pasta.isDirectory()){

            // lista de todos os arquivos dentro da pasta (aqui ele pega os arquivos dentro da pasta e salva na lista)
            File[] arquivos = pasta.listFiles();

            // cria o builder
            DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
            DocumentBuilder builder = factory.newDocumentBuilder();

            // verifica se a pasta não está vazia
            if(arquivos != null){
                System.out.println("LISTA DE ARQUIVOS: ");
                
                // percorre todos os arquivos da pasta
                for (File arquivo : arquivos){
                    lerXML(arquivo, builder); // chama a função lerXML para cada arquivo
                }
            }
            else {
                System.out.println("A pasta está vazia.");
                return;
            }
        }
        
        else{
            System.out.println("A pasta não existe ou não é um diretório");
            return;
        }
    }



    // ESSA FUNÇÃO FAZ LER ARQUIVOS EM EXCEL
    private static void lerModelos(){

        // cria o file com o caminho do arquivo
        File arquivo = new File(caminho_arquivo);
        
        // verifica se o arquivo existe
        if(!arquivo.exists()){
            System.out.println("O arquivo não existe.");
            return;
        }

        // Abre o fluxo de entrada de arquivo dentro do try-with-resources (fecha automaticamente ao terminar)
        try (FileInputStream fis = new FileInputStream(arquivo);
             Workbook workbook = WorkbookFactory.create(fis)) {

            // acessa a primeira aba/planilha (índice 0)
            Sheet aba = workbook.getSheetAt(0);
            System.out.println("Nome da aba: " + aba.getSheetName());
            System.out.println("--------------------------------------------------");

            // DataFormatter serve para que valores numéricos e datas sejam convertidos corretamente para texto
            DataFormatter formatter = new DataFormatter();

            // percorre todas as linhas da aba
            for (Row linha : aba) {
                // Ignora a linha de cabeçalho (linha 0)
                if (linha.getRowNum() == 0) continue;

                // Cria uma celula e pega a coluna (começa em 0)
                Cell celulaItem = linha.getCell(0);
                Cell celulaCategoria = linha.getCell(2);

                // Pega o valor da celula
                String item = formatter.formatCellValue(celulaItem);
                String categoria = formatter.formatCellValue(celulaCategoria);

                System.out.println("Código: " + item + " | Setor: " + categoria);
            }
        }
        catch (Exception e) {
            System.err.println("Erro ao ler o arquivo Excel: " + e.getMessage());
            e.printStackTrace();
        }
    }




    // função que vai ler o xml e retornar os codigos da referencia
    private static void lerXML(File arquivo, DocumentBuilder builder){
        try{
            // pega o arquivo e parse ele
            Document doc = builder.parse(arquivo);
            doc.getDocumentElement().normalize();

            // pega os "itens" que estão dentro da tag "ITEM"
            NodeList item = doc.getElementsByTagName("ITEM");
            
            // percorre todos os "itens"
            for (int i = 0; i < item.getLength(); i++){
                
                // aqui ele ta pegando o item
                Node node = item.item(i);
                
                // verifica se o item é um elemento
                if (node.getNodeType() == Node.ELEMENT_NODE){
                    // transforma o node em elemento
                    Element elemento = (Element) node;
                    
                    // cria uma string chamada referencia e pega o atributo REFERENCIA do item
                    String referencia =  elemento.getAttribute("REFERENCIA");
                    
                    System.out.println("Achei esse código: " + referencia);
                }

            }

        }
        catch (Exception e){
            e.printStackTrace();
        }
    }



    */
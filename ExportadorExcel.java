import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExportadorExcel {

    /* Verifica se já existe esse pedido no xlsx */
    public static Set<String> pedidosExistentes(String caminho){

        Set<String> pedidos = new HashSet<>(); // Cria um set para guardar os pedidos
        File arquivo = new File(caminho); // Cria o arquivo
        

        if(!arquivo.exists() || !arquivo.isFile()) return pedidos; // Se o arquivo não existir ou não for um arquivo, retorna o set vazio

        try(FileInputStream file = new FileInputStream(arquivo); Workbook workbook = WorkbookFactory.create(file)){

            Sheet aba = workbook.getSheetAt(0); // Pega a primeira aba
            DataFormatter formatter = new DataFormatter();


            for(Row linha : aba){
                if(linha.getRowNum() == 0) continue; // Pula o cabeçalho

                Cell celulaPedido = linha.getCell(0);// Pega a coluna 0 (Pedido)

                if(celulaPedido != null){ // Verifica se a célula não está vazia
                    String pedido = formatter.formatCellValue(celulaPedido).trim(); // Pega o valor da coluna 0 e remove espaços extras
                    
                    if(!pedido.isEmpty()) { // Verifica se o pedido não está vazio
                        pedidos.add(pedido); // Adiciona o pedido no set
                    }
                }
            }
            
        }
        catch(Exception e){
            System.err.println("Erro ao ler o arquivo Excel: " + e.getMessage());
        }

        return pedidos;
    }



    public static class DadosGrupo {
        private int quantidade;
        private double valorTotal;

        public DadosGrupo() {
            this.quantidade = 0;
            this.valorTotal = 0.0;
        }

        public DadosGrupo(int quantidade, double valorTotal) {
            this.quantidade = quantidade;
            this.valorTotal = valorTotal;
        }

        public void adicionar(int qtd, double valor) {
            this.quantidade += qtd;
            this.valorTotal += valor;
        }

        public int getQuantidade() {
            return quantidade;
        }

        public double getValorTotal() {
            return valorTotal;
        }
    }

    public static boolean exportar(Map<String, Map<String, DadosGrupo>> dadosPorArquivo, String caminho){

        File arquivo = new File(caminho);
        Workbook workbook = null;
        Sheet aba = null;

        try{ // cria um workbook para ser editado

            boolean arquivoExiste = arquivo.exists();

            if (arquivoExiste) {

                // Abre o arquivo existente para modificação
                try (FileInputStream fis = new FileInputStream(arquivo)) {
                    workbook = WorkbookFactory.create(fis); // cria o workbook a partir do arquivo existente
                }

                aba = workbook.getSheetAt(0); // pega a primeira aba

                // Atualiza o cabeçalho se a coluna de valor ainda não existir
                Row cabecalhoExistente = aba.getRow(0);
                if (cabecalhoExistente != null && cabecalhoExistente.getCell(3) == null) {
                    Cell celulaValor = cabecalhoExistente.createCell(3);
                    celulaValor.setCellValue("Valor Total");
                    CellStyle estiloCabecalho = workbook.createCellStyle();
                    Font fonte = workbook.createFont();
                    fonte.setBold(true);
                    estiloCabecalho.setFont(fonte);
                    celulaValor.setCellStyle(estiloCabecalho);
                }
            } 
            else {

                workbook = new XSSFWorkbook(); // cria um novo workbook
                aba = workbook.createSheet("Planilha"); // cria a planilha
                
                Row cabecalho = aba.createRow(0); // cria cabaçalho na primeira linha


                /* Cria as celulas do cabecalho*/
                Cell cabecalhoPedido = cabecalho.createCell(0);
                Cell cabecalhoGrupo = cabecalho.createCell(1);
                Cell cabecalhoQuantidade = cabecalho.createCell(2);
                Cell cabecalhoValor = cabecalho.createCell(3);


                /* Coloca valores no cabecalho*/
                cabecalhoPedido.setCellValue("Pedido");
                cabecalhoGrupo.setCellValue("Grupo");
                cabecalhoQuantidade.setCellValue("Quantidade");
                cabecalhoValor.setCellValue("Valor Total");


                /* deixa o cabeçalho em negrito*/
                CellStyle estiloCabecalho = workbook.createCellStyle(); // cria o estilo do cabecalho
                Font fonte = workbook.createFont(); // cria a fonte
                fonte.setBold(true); // define a fonte como negrito
                estiloCabecalho.setFont(fonte); // aplica a fonte no estilo do cabecalho
                cabecalhoPedido.setCellStyle(estiloCabecalho);
                cabecalhoGrupo.setCellStyle(estiloCabecalho); // aplica o estilo do cabecalho na coluna grupo
                cabecalhoQuantidade.setCellStyle(estiloCabecalho); // aplica o estilo do cabecalho na coluna quantidade
                cabecalhoValor.setCellStyle(estiloCabecalho);
            }

            // Estilo numérico/moeda para o valor total
            CellStyle estiloMoeda = workbook.createCellStyle();
            estiloMoeda.setDataFormat(workbook.createDataFormat().getFormat("R$ #,##0.00"));

            /* Coloca os valores na planilha */
            int numeroLinha = arquivoExiste ? aba.getLastRowNum() + 1 : 1; // se ja existe, continua na proxima linha
            int totalLinha = 0;

            for (Map.Entry<String, Map<String, DadosGrupo>> arquivoEntry : dadosPorArquivo.entrySet()){ 

                String nomeArquivo = arquivoEntry.getKey(); // recebe o nome do arquivo
                String nomePedido = nomeArquivo.length() >= 5 ? nomeArquivo.substring(0, 5) : nomeArquivo;

                Map<String, DadosGrupo> grupos = arquivoEntry.getValue(); // recebe os grupos do arquivo, quantidade e valor

                // Percorre cada grupo do arquivo atual
                for (Map.Entry<String, DadosGrupo> grupoEntry : grupos.entrySet()){
                    Row linha = aba.createRow(numeroLinha++); // cria uma nova linha
                    DadosGrupo dadosGrupo = grupoEntry.getValue();
                    
                    linha.createCell(0).setCellValue(nomePedido); // Pedido
                    linha.createCell(1).setCellValue(grupoEntry.getKey()); // Grupo
                    linha.createCell(2).setCellValue(dadosGrupo.getQuantidade()); // Quantidade
                    
                    Cell celulaValor = linha.createCell(3);
                    celulaValor.setCellValue(dadosGrupo.getValorTotal()); // Valor Total
                    celulaValor.setCellStyle(estiloMoeda);

                    totalLinha++;
                }
            }

            aba.autoSizeColumn(0); // ajusta a largura automaticamente da coluna 0
            aba.autoSizeColumn(1); // ajusta a largura automaticamente da coluna 1
            aba.autoSizeColumn(2); // ajusta a largura automaticamente da coluna 2
            aba.autoSizeColumn(3); // ajusta a largura automaticamente da coluna 3


            /* cria o arquivo e escreve o workbook no arquivo */
            try(FileOutputStream file = new FileOutputStream(caminho)){ // cria o arquivo
                workbook.write(file); // escreve no arquivo
                System.out.println("Excel gerado com sucesso em: " + caminho); 
                return true;
            }
        }

        catch(Exception e){
            System.err.println("Erro ao exportar para Excel: " + e.getMessage());
            return false;
        }
        
        /* Usamos o finally para garantir que o arquivo seja fechado, mesmo se ocorrer um erro */
        finally{
            if(workbook != null){
                try{
                    workbook.close();
                }
                catch(Exception e){
                    System.err.println("Erro ao fechar o workbook: " + e.getMessage());
                }
            }
        }

    }
}

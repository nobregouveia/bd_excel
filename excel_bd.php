<?php

header("Content-type: text/php charset=utf-8");
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PDO;
use PDOException;

class Database {
    private $host = 'localhost:3306';
    private $db = 'bd_excel';
    private $user = 'root';
    private $pass = 'root';
    private $charset = 'utf8mb4';
    private $pdo;

    public function __construct() {
        $dsn = "mysql:host=$this->host;dbname=$this->db;charset=$this->charset";
        $options = [
            PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
            PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC,
            PDO::ATTR_EMULATE_PREPARES => false,
        ];

        try {
            $this->pdo = new PDO($dsn, $this->user, $this->pass, $options);
        } catch (PDOException $e) {
            throw new PDOException($e->getMessage(), (int)$e->getCode());
        }
    }

    public function getPdo() {
        return $this->pdo;
    }
}

class ExcelReader {
    public static function read($filePath) {
        $spreadsheet = IOFactory::load($filePath);
        $sheet = $spreadsheet->getActiveSheet();
        $data = [];
        $firstRow = true;
        foreach ($sheet->getRowIterator() as $row) {
            if ($firstRow) {
                $firstRow = false;
                continue; // Pula a primeira linha (cabeçalho)
            }
            $rowData = [];
            foreach ($row->getCellIterator() as $cell) {
                $rowData[] = $cell->getFormattedValue();
            }
            $data[] = $rowData;
        }
        return $data;
    }
}

class Fornecedor {
    private $pdo;

    public function __construct($pdo) {
        $this->pdo = $pdo;
    }

    public function insert($data) {
        try {
            $stmt = $this->pdo->prepare("INSERT INTO bd_excel.fornecedor (nome_fornecedor, cnpj_cpf, data_homologacao, situacao, escopo_servicos, qualificado, pessoa_f_j, contato, telefone, email, banco, agencia, conta, fcpa) 
                                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
            $stmt->execute($data);
            return $this->pdo->lastInsertId();
        } catch (PDOException $e) {
            var_dump("Erro ao inserir fornecedor: " . $e->getMessage());
            return false;
        }
    }

    public function getByCnpj($cnpj) {
        $stmt = $this->pdo->prepare("SELECT id FROM bd_excel.fornecedor WHERE cnpj_cpf = ?");
        $stmt->execute([$cnpj]);
        return $stmt->fetchColumn();
    }
}

class Nota {
    private $pdo;

    public function __construct($pdo) {
        $this->pdo = $pdo;
    }

    public function insert($data, $fornecedorId) {
        $stmt = $this->pdo->prepare("INSERT INTO bd_excel.notas (fornecedor_id, n1, n2, n3, n4, n5, n6, n7, n8, n9, n10) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)");
        array_unshift($data, $fornecedorId);

        foreach ($data as &$value) {
            if ($value === '' || $value === null) {
                $value = 0; // Ou pode substituir por 0 se preferir
            }
        }

        if (count($data) !== 11) {
            var_dump("Número incorreto de parâmetros para a consulta SQL.");
            return false;
        }

        try {
            $stmt->execute($data);
            var_dump("Notas inseridas com sucesso.");
            return true;
        } catch (PDOException $e) {
            var_dump("Erro ao inserir notas: " . $e->getMessage());
            return false;
        }
    }
}

class FileUploader {
    private $allowedExtensions = ['xls', 'xlsx', 'xlsm'];
    private $uploadDir = './uploads/';

    public function upload($file) {
        $fileTmpPath = $file['tmp_name'];
        $fileName = $file['name'];
        $fileNameCmps = explode(".", $fileName);
        $fileExtension = strtolower(end($fileNameCmps));

        if (in_array($fileExtension, $this->allowedExtensions)) {
            $destPath = $this->uploadDir . $fileName;
            if (move_uploaded_file($fileTmpPath, $destPath)) {
                return $destPath;
            } else {
                throw new Exception("Erro ao mover o arquivo para o diretório de upload.");
            }
        } else {
            throw new Exception("Upload falhou. Somente arquivos Excel são permitidos.");
        }
    }
}

class ExcelProcessor {
    private $pdo;

    public function __construct($pdo) {
        $this->pdo = $pdo;
    }

    public function process($filePath) {
        $data = ExcelReader::read($filePath);
        $fornecedor = new Fornecedor($this->pdo);
        $nota = new Nota($this->pdo);

        foreach ($data as $row) {
            $cnpj = $row[1];
            if (empty($cnpj)) {
                continue; // Ignorar linhas sem CNPJ
            }

            $fornecedorId = $fornecedor->getByCnpj($cnpj);
            if ($fornecedorId === false) {
                $fornecedorData = array_merge(
                    array_slice($row, 0, 5),    // Colunas A-E
                    array_slice($row, 16, 9)   // Colunas Q-Y
                );
                $fornecedorId = $fornecedor->insert($fornecedorData);
            }

            if ($fornecedorId !== false) {
                $notasData = array_slice($row, 5, 10);
                $nota->insert($notasData, $fornecedorId);
            }
        }
    }
}

if ($_SERVER['REQUEST_METHOD'] == 'POST' && isset($_FILES['file'])) {
    try {
        $uploader = new FileUploader();
        $filePath = $uploader->upload($_FILES['file']);

        $database = new Database();
        $pdo = $database->getPdo();

        $processor = new ExcelProcessor($pdo);
        $processor->process($filePath);

        echo "Dados inseridos com sucesso!";
    } catch (Exception $e) {
        echo $e->getMessage();
    }
} else {
    echo "Nenhum arquivo enviado.";
}

?>
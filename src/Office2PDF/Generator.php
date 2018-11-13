<?php

namespace Office2PDF;

use Exception;

class Generator 
{
    /**
     * Supported formats
     */
    const FORMATS = ['dwg', 'doc', 'docx', 'wps', 'pdf', 'xlsx', 'xls', 'ppt', 'pptx'];

    /**
     * A list of files to convert 
     *
     * @var array
     */
    private $fileNames = [];

    /**
     * Path of output files 
     *
     * @var string
     */
    private $outputDirectory;
    
    /**
     * Construct and stt file(s) to convert into PDF
     *
     * @param array $fileNames
     */
    public function __construct(array $fileNames = []) 
    {
        $this->checkRequirements();

        if (count($fileNames) > 0) {
            $this->addFiles($fileNames);
        }
    }
    
    /**
     * This checks program requirements 
     *
     * @return bool
     */
    public function checkRequirements(): bool
    {
        // Check OS 
        if(PHP_OS != 'Linux') {
            throw new Exception('Operating system is not supported. The program only runs on Linux server.');
            return false;
        }
    
        // Check if java has been installed
        exec('java -version > NULL && echo yes || echo no', $output);
    
        if($output[0] == 'no') {
            throw new Exception('There is no Java environment.');
            return false;
        }
        
        // Check if liberoffice has been installed
        exec('libreoffice --version > NULL && echo yes || echo no', $output);
        if($output[0] == 'no') {
            throw new Exception('LibreOffice has not been installed.');
            return false;
        }
      
        return true;
    }

    /**
     * Check if a given file is supported
     *
     * @param string $fileName
     * @return boolean
     */
    public function isSupported(string $fileName): bool
    {
        if (!is_file($fileName)) {
            throw new Exception('The file ' . $fileName . ' does not exist.');
            return false;
        }

        $fileInfo   = pathinfo($fileName);
        $fileFormat = $fileInfo['extension'];
        
        return in_array($fileFormat, self::FORMATS);
    }

    /**
     * Specify an output directory to store converted files
     *
     * @param string $path
     * @return void
     */
    public function setOutputDir(string $path)
    {
        if (!is_dir($path)) {
            if (!mkdir($path)) {
                throw new Exception('Cannot create path "' . $path .'".');
                return;
            }
        }

        $this->outputDirectory = (substr($path, -1) !== '/') ? $path . '/' : $path;
    }

    /**
     * Add one file to be converted
     *
     * @param string $fileName
     * @return void
     */
    public function addFile(string $fileName)
    {
        if($this->isSupported($fileName)) {
            $this->fileNames[] = $fileName;
            return;
        }

        throw new Exception('Failed to add "' . basename($fileName) . '". The format is not supported.');
    }
    
    /**
     * Add a bunch of files to be converted
     *
     * @param array $fileNames
     * @return void
     */
    public function addFiles(array $fileNames)
    {
        foreach ($fileNames as $file) {
            $this->addFile($file);
        }
    }

    /**
     * Converts given files to PDF and stores them into $outputDirectoy
     *
     * @param string $destination
     * @return array output stat 
     */
    public function convert(string $destination = ''): array
    {
        if(!empty(trim($destination))) {
            $this->setOutputDir($destination);
        }
        
        if(!$this->outputDirectory) {
            throw new Exception('No output directory has been defined.');
        }

        $successCount = 0;
        $failedCount  = 0;
        $converted    = [];
        
        if (is_dir($this->outputDirectory) === false) {
            throw new Exception('Directory "' . $this->outputDirectory . '" does not exist.');
        }

        foreach($this->fileNames as $file) {
            if (!file_exists($file)) {
                throw new Exception('File "' . $file . '" does not exist.');
            }

            $fileParts   = pathinfo($file);
            $converted[] = $this->outputDirectory . $fileParts['filename'] . '.pdf';

            try {
                $command = 'libreoffice --invisible --convert-to pdf '.$file.' --outdir '.$this->outputDirectory;
                $cmdMsg  = passthru($command);
                
                $successCount++;
            } catch (Exception $e) {
                $failedCount ++;
            }
        }

        return [
            'success'        => $successCount,
            'failed'         => $failedCount,
            'convertedFiles' => $converted 
        ];
    }
}
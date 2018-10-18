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
    public function __construct(array $fileNames) 
    {
        $this->fileNames = $fileNames;
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
        $output = passthru('java -version && echo java');
        if($output != 'java') {
            throw new Exception('There is no Java environment.');
            return false;
        }
        
        // Check if liberoffice has been installed
        $output = passthru('libreoffice --version && echo libreoffice');
        if($output != 'libreoffice') {
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
    public function isSupported(string $fileName)
    {
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
     * @return integer Number of converted files 
     */
    public function convert(string $destination): int
    {
        $this->outputDirectory = !is_null($destination) ? $destination : $this->outputDirectory;

        $this->setOutputDir($this->outputDirectory);

        $successCount = 0;
        
        if (is_dir($this->outputDirectory) === false) {
            throw new Exception('Directory "' . $this->outputDirectory . '" does not exist.');
            return;
        }

        foreach($this->fileNames as $file) {
            if (!file_exists($file)) {
                throw new Exception('File "' . $file . '" does not exist.');
                return;
            }

            $fileParts = pathinfo($file);
            $output    = $this-> outputDirectory . $fileParts['filename'] . '.pdf';

            try {
                $command = 'libreoffice --invisible --convert-to pdf '.$file.' --outdir '.$this->outputDirectory;
                $cmdMsg  = passthru($command);
                
                $successCount++;
            } catch (Exception $e) {
                return 0;
            }
        }

        return $successCount;
    }
}
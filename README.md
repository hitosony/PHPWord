# PHPWord

# Installasi PHPWord
1. Install Composer
2. Buka CMD, arahkan ke folder aplikasi "xampp/htdocs/{nama_folder}
3. ketik : composer require phpoffice/phpword

# Configurasi Codeigniter
1. Buka file config/config.php, cari baris
   - $config['composer_autoload'] = '';
   
   ubah menjadi
   - $config['composer_autoload'] = 'vendor/autoload.php';
2. Buka file controller lalu tambahkan script ini diatas baris class
   - use PhpOffice\PhpWord\TemplateProcessor;

# Contoh function
    
    public function printFormulirRegis($id)
    {
    
        $location = base_url('word/');
        $templateProcessor = new TemplateProcessor($location . 'Permohonan_Hak_Pakai.docx');
        $row = $this->Pemilik_model->getData($id);
        
        $nama = $row->nama;
        $id_pasar = $row->id_pasar;
        $id_lokasi = $row->id_lokasi;
        $blok = $row->blok;
        $no_lokasi = $row->no_lokasi;
        $id_gender = $row->id_gender;
        $alamat = $row->alamat;
        $produk = $row->produk;
        $tempat_lahir = $row->tempat_lahir;
        $tgl_lahir = format_indo($row->tgl_lahir);
        $date_created = format_indo(date('Y-m-d'));

        $filename = 'Formulir Permohonan ' . $nama . '.docx';

        $templateProcessor->setValue('nama', $nama);
        $templateProcessor->setValue('id_pasar', $id_pasar);
        $templateProcessor->setValue('id_lokasi', $id_lokasi);
        $templateProcessor->setValue('blok', $blok);
        $templateProcessor->setValue('no_lokasi', $no_lokasi);
        $templateProcessor->setValue('id_gender', $id_gender);
        $templateProcessor->setValue('produk', $produk);
        $templateProcessor->setValue('alamat', $alamat);
        $templateProcessor->setValue('tempat_lahir', $tempat_lahir);
        $templateProcessor->setValue('tgl_lahir', $tgl_lahir);
        $templateProcessor->setValue('date_created', $date_created);
        
        $templateProcessor->saveAs($filename);
        
        header('Content-Description: File Transfer');
        header('Content-Type: application/octet-stream');
        header('Content-Disposition: attachment; filename=' . $filename);
        header('Content-Transfer-Encoding: binary');
        header('Expires: 0');
        header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
        header('Pragma: public');
        header('Content-Length: ' . filesize($filename));
        flush();
        readfile($filename);
        unlink($filename); // deletes the temporary file
        exit;
        redirect(site_url('admin/pemilik'));
    }

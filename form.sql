CREATE TABLE history_file (
  id INT AUTO_INCREMENT PRIMARY KEY,
  nama_file VARCHAR(255),
  jenis_form ENUM('perjadin', 'pendataan'),
  tanggal_generate DATETIME
);

<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
Route::get('gen_db_pg', function() {	    
    $tableNames = config('dict.table_ignore'); // Loại bỏ những table không cần export

// Chuyển đổi mảng thành chuỗi các tên bảng được phân cách bởi dấu phẩy và bao quanh bởi dấu nháy đơn
    $tableNamesString = "'" . implode("','", $tableNames) . "'";
    $sql_table_str = "
                SELECT z.table_name, z2.column_name, z2.data_type, z2.column_default, z2.is_nullable, pg_catalog.col_description(pgc.oid, z2.ordinal_position::int) as comment
		    FROM information_schema.tables z
				JOIN information_schema.columns z2 on z.table_name = z2.table_name  
                JOIN pg_catalog.pg_class pgc ON z.table_name = pgc.relname 
		WHERE table_type = 'BASE TABLE' 
		    AND z.table_schema NOT IN ('pg_catalog', 'information_schema')
            AND NOT pgc.relispartition    -- exclude child partitions -- add by Tuanhv-27-05-2022
            AND z.table_name NOT IN ($tableNamesString)
		order by z.table_name, z2.ordinal_position
    ";
	
    $data_tables = DB::select($sql_table_str);	   

    $sql_get_pri_unique = "
        with t1 as (
		select kcu.table_schema,
		       kcu.table_name,
		       tco.constraint_name,
		       kcu.ordinal_position as position,
		       kcu.column_name,
					 tco.constraint_type
		from information_schema.table_constraints tco
		join information_schema.key_column_usage kcu 
		     on kcu.constraint_name = tco.constraint_name
		     and kcu.constraint_schema = tco.constraint_schema
		     and kcu.constraint_name = tco.constraint_name
		where tco.constraint_type in ('FOREIGN KEY', 'PRIMARY KEY', 'UNIQUE')
        and kcu.table_name not in ($tableNamesString)
		order by kcu.table_schema,
		         kcu.table_name,
		         position
		) select table_name, constraint_name as index_name, string_agg(t1.column_name, ',') as columns, constraint_type
		from t1 group by table_name, constraint_name,constraint_type
    ";
    
    $data_indexs = \DB::select($sql_get_pri_unique);
    // how to fix $data_indexs to get all table name    

    $tables = [];
    $datas = [];
    $no = 1;
    $fields_not_comment = [];    
    foreach ($data_tables as $dkey => $data_table) {
        // if ($data_table->comment) {
        //     $namejp = $data_table->comment;
        // }else{
            if(array_key_exists($data_table->column_name, config('dict.sakin'))) {
                $namejp = config('dict.sakin')[$data_table->column_name];
            } else {
                $namejp = ucwords(str_replace('_', " ", $data_table->column_name)) ;
                $fields_not_comment[$data_table->table_name][] = $data_table->column_name;
            }
        // }        
        // $table_name = strtolower(substr($data_table->table_name, 0, 30));
        $table_name = strtolower($data_table->table_name);                               
        $datas[$table_name][$dkey]['no'] = $dkey;
        $datas[$table_name][$dkey]['name_jp'] = $namejp;
        $datas[$table_name][$dkey]['column_name'] = $data_table->column_name;
        $datas[$table_name][$dkey]['data_type'] = $data_table->data_type;
        $datas[$table_name][$dkey]['not_null'] = $data_table->is_nullable == 'YES' ? '':'Yes';
        $datas[$table_name][$dkey]['column_default'] = $data_table->column_default ?? '';
        $datas[$table_name][$dkey]['comment'] = $data_table->comment ?? '';
        $tables[$table_name] = $table_name;
    }        
    foreach ($fields_not_comment as $key_table => $item) {
        $list_error = '';
        foreach ($item as $key => $value) {
            $list_error .= $value. ' ';
            //Debug::log('column_needed_translation', $key_table. '---' . $value . '---' .ucwords(str_replace('_', " ", $value)));
        }
    }

    foreach ($datas as $d_key => $data) {
        $no = 1;
        foreach ($data as $dkey => $dvalue) {
            $datas[$d_key][$dkey]['no'] = $no++;
        }
    }

    $data_frs = [];
    foreach ($data_indexs as $key => $data_index) {
        $table = $data_index->table_name;
        if ($data_index->constraint_type == 'FOREIGN KEY') {
            if (!isset($data_frs[$table])) {
                $no = 1;
            }
            $data_frs[$table][$key]['no'] = $no++;
            $data_frs[$table][$key]['name'] = $data_index->index_name;
            $data_frs[$table][$key]['column'] = $data_index->columns;
            $data_frs[$table][$key]['tbl_re'] = '';
            $data_frs[$table][$key]['tbl_re_col'] = '';
        }
    }

    $datas_pri_uni = [];
    $current_table = '';
    foreach ($data_indexs as $kd => $data_index) {
        $tbl = $data_index->table_name;
        if ($data_index->constraint_type == 'PRIMARY KEY' || $data_index->constraint_type == 'UNIQUE') {
            $datas_pri_uni[$tbl][$kd]['no'] = $kd;
            $datas_pri_uni[$tbl][$kd]['name'] = $data_index->index_name;
            $datas_pri_uni[$tbl][$kd]['column'] = $data_index->columns;
            $datas_pri_uni[$tbl][$kd]['p_type'] = $data_index->constraint_type == 'PRIMARY KEY' ? 'Yes':'';
            $datas_pri_uni[$tbl][$kd]['u_type'] = $data_index->constraint_type == 'UNIQUE' ? 'Yes':'';
        }
    }

    foreach ($datas_pri_uni as $dp_k => $dp) {
        $no = 1;
        foreach ($dp as $dpkey => $value) {
            $datas_pri_uni[$dp_k][$dpkey]['no'] = $no++;
        }
    }

    //template file
    $file = storage_path('db_cyinder_gatelock.xlsx');
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file);
    // Add export all table TuanHa 2022.05.25
    $spreadsheet->setActiveSheetIndex(0);
    $i = 2; //  bắt đầu ghi vào tên các table từ dòng thứ 2
    foreach( $tables as $table) {
        $spreadsheet->getActiveSheet()->SetCellValue('A'.$i, ($i-1));
        $spreadsheet->getActiveSheet()->SetCellValue('B'.$i, $table);
        $spreadsheet->getActiveSheet()->getCell('B'.$i)->getHyperlink()->setUrl("sheet://'" .  substr($table, 0, 30) . "'!A1"); // gắn hyperlink từ bảng chính tới sheet table.
        // Định dạng màu xanh cho hyperlink
        $spreadsheet->getActiveSheet()->getStyle('B'.$i)->applyFromArray([
            'font' => [
                'color' => ['rgb' => '0000FF'], // Màu xanh
                'underline' => 'single' // Gạch chân để giống hyperlink
            ]
        ]);
        $i++;
    }

    $spreadsheet->getActiveSheet()->getStyle('A1:B'.($i-1))->getBorders()
        ->getAllBorders()
        ->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
    // End export all table TuanHa 2022.05.25
    foreach ($tables as $table) {        
        $clonedWorksheet = clone $spreadsheet->getSheetByName('template');
        $clonedWorksheet->setTitle(substr($table, 0, 30));
        $spreadsheet->addSheet($clonedWorksheet);
        // Những bảng có ký tự 'yoyaku' hoặc 'zimmer' thì tô màu sheet đó lên.
        if(strpos($table, "yoyaku") !== FALSE || strpos($table, "zimmer") !== FALSE) {
            $spreadsheet->setActiveSheetIndexByName(substr($table, 0, 30))->getTabColor()->setARGB('FF0000');
        } else {
            $spreadsheet->setActiveSheetIndexByName(substr($table, 0, 30));
        }

        $worksheet = $spreadsheet->getActiveSheet();
        $logic_final = $table;
        if (strpos($logic_final, 't_') === 0) {
            $logic_final = substr($logic_final, 2);
        }

        // Kiểm tra và loại bỏ 'c_' nếu có
        if (strpos($logic_final, 'm_') === 0) {
            $logic_final = substr($logic_final, 2);
        }
        $logic_name = $logic_final;        
        $worksheet->getCell('C5')->setValue(ucwords(str_replace('_', " ", $logic_name)));
        $worksheet->getCell('C6')->setValue($table);
        $worksheet->getCell('F5')->setValue('PgSQL');

        $data_struct = $datas[$table];
        $data_p = [];
        if (isset($datas_pri_uni[$table])) {
            $data_p = $datas_pri_uni[$table];
        }
        $data_fa = [];
        if (isset($data_frs[$table])) {
            $data_fa = $data_frs[$table];
        }		
        foreach ($worksheet->getRowIterator() as $k_row => $row) {
            foreach( $row->getCellIterator() as $k_col => $cell ){
                $value = $cell->getCalculatedValue();
                if ($value == 'カラム情報') {
                    $new_row = $k_row + 2;
                    foreach ($data_struct as $data_s) {
                        $worksheet->insertNewRowBefore($new_row, 1);
                        // $spreadsheet->getActiveSheet()->getStyle('A'.$new_row)->getFill()
                        //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                        //     ->getStartColor()->setARGB('ffffff');
                        // $spreadsheet->getActiveSheet()->getStyle('B'.$new_row)->getFill()
                        //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                        //     ->getStartColor()->setARGB('ffffff');
                        // $spreadsheet->getActiveSheet()->getStyle('C'.$new_row)->getFill()
                        //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                        //     ->getStartColor()->setARGB('ffffff');
                        // $spreadsheet->getActiveSheet()->getStyle('D'.$new_row)->getFill()
                        //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                        //     ->getStartColor()->setARGB('ffffff');
                        // $spreadsheet->getActiveSheet()->getStyle('E'.$new_row)->getFill()
                        //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                        //     ->getStartColor()->setARGB('ffffff');
                        // $spreadsheet->getActiveSheet()->getStyle('F'.$new_row)->getFill()
                        //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                        //     ->getStartColor()->setARGB('ffffff');
                        // $spreadsheet->getActiveSheet()->getStyle('G'.$new_row)->getFill()
                        //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                        //     ->getStartColor()->setARGB('ffffff');
                                   // Thiết lập màu nền cho các ô từ cột A đến G
                $columns = range('A', 'G');
                foreach ($columns as $column) {
                    $spreadsheet->getActiveSheet()->getStyle($column.$new_row)->getFill()
                        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                        ->getStartColor()->setARGB('ffffff');
                    // Bỏ in đậm cho các ô từ cột A đến G
                    $spreadsheet->getActiveSheet()->getStyle($column.$new_row)->getFont()->setBold(false);
                }
                        $worksheet->getCell('A'.$new_row)->setValue($data_s['no']);
                        $worksheet->getCell('B'.$new_row)->setValue($data_s['name_jp']);
                        $worksheet->getCell('C'.$new_row)->setValue($data_s['column_name']);
                        $worksheet->getCell('D'.$new_row)->setValue($data_s['data_type']);
                        $worksheet->getCell('E'.$new_row)->setValue($data_s['not_null']);
                        $worksheet->getCell('F'.$new_row)->setValue($data_s['column_default']);
                        $worksheet->getCell('G'.$new_row)->setValue($data_s['comment']);
                        $new_row++;
                    }
                }
                break;
            }
        }

        if (count($data_p)) {
            foreach ($worksheet->getRowIterator() as $k_row => $row) {
                foreach( $row->getCellIterator() as $k_col => $cell ){
                    $value = $cell->getCalculatedValue();
                    if ($value == 'インデックス情報') {
                        $new_row = $k_row + 2;
                        foreach ($data_p as $dp) {
                            $worksheet->insertNewRowBefore($new_row, 1);
                            // $spreadsheet->getActiveSheet()->getStyle('A'.$new_row)->getFill()
                            //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            //     ->getStartColor()->setARGB('ffffff');
                            // $spreadsheet->getActiveSheet()->getStyle('B'.$new_row)->getFill()
                            //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            //     ->getStartColor()->setARGB('ffffff');
                            // $spreadsheet->getActiveSheet()->getStyle('C'.$new_row)->getFill()
                            //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            //     ->getStartColor()->setARGB('ffffff');
                            // $spreadsheet->getActiveSheet()->getStyle('D'.$new_row)->getFill()
                            //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            //     ->getStartColor()->setARGB('ffffff');
                            // $spreadsheet->getActiveSheet()->getStyle('E'.$new_row)->getFill()
                            //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            //     ->getStartColor()->setARGB('ffffff');
                            // $spreadsheet->getActiveSheet()->getStyle('F'.$new_row)->getFill()
                            //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            //     ->getStartColor()->setARGB('ffffff');
                            // $spreadsheet->getActiveSheet()->getStyle('G'.$new_row)->getFill()
                            //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            //     ->getStartColor()->setARGB('ffffff');
                            $columns = range('A', 'G');
                            foreach ($columns as $column) {
                                $spreadsheet->getActiveSheet()->getStyle($column.$new_row)->getFill()
                                    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                    ->getStartColor()->setARGB('ffffff');
                                // Bỏ in đậm cho các ô từ cột A đến G
                                $spreadsheet->getActiveSheet()->getStyle($column.$new_row)->getFont()->setBold(false);
                            }
                            $worksheet->getCell('A'.$new_row)->setValue($dp['no']);
                            $worksheet->getCell('B'.$new_row)->setValue($dp['name']);
                            $worksheet->getCell('C'.$new_row)->setValue($dp['column']);
                            $worksheet->getCell('E'.$new_row)->setValue($dp['p_type']);
                            $worksheet->getCell('F'.$new_row)->setValue($dp['u_type']);
                            $new_row++;
                        }
                    }
                    break;
                }
            }
        }

        try {
            if (count($data_fa)) {
                foreach ($worksheet->getRowIterator() as $k_row => $row) {
                    foreach( $row->getCellIterator() as $k_col => $cell ){
                        $value = $cell->getCalculatedValue();
                        if ($value == '外部キー情報') {
                            $new_row = $k_row + 2;
                            foreach ($data_fa as $df) {
                                $worksheet->insertNewRowBefore($new_row, 1);
                                // $spreadsheet->getActiveSheet()->getStyle('A'.$new_row)->getFill()
                                //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                //     ->getStartColor()->setARGB('ffffff');
                                // $spreadsheet->getActiveSheet()->getStyle('B'.$new_row)->getFill()
                                //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                //     ->getStartColor()->setARGB('ffffff');
                                // $spreadsheet->getActiveSheet()->getStyle('C'.$new_row)->getFill()
                                //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                //     ->getStartColor()->setARGB('ffffff');
                                // $spreadsheet->getActiveSheet()->getStyle('D'.$new_row)->getFill()
                                //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                //     ->getStartColor()->setARGB('ffffff');
                                // $spreadsheet->getActiveSheet()->getStyle('E'.$new_row)->getFill()
                                //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                //     ->getStartColor()->setARGB('ffffff');
                                // $spreadsheet->getActiveSheet()->getStyle('F'.$new_row)->getFill()
                                //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                //     ->getStartColor()->setARGB('ffffff');
                                // $spreadsheet->getActiveSheet()->getStyle('G'.$new_row)->getFill()
                                //     ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                //     ->getStartColor()->setARGB('ffffff');
                                $columns = range('A', 'G');
                                foreach ($columns as $column) {
                                    $spreadsheet->getActiveSheet()->getStyle($column.$new_row)->getFill()
                                        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                        ->getStartColor()->setARGB('ffffff');
                                    // Bỏ in đậm cho các ô từ cột A đến G
                                    $spreadsheet->getActiveSheet()->getStyle($column.$new_row)->getFont()->setBold(false);
                                }
                                $worksheet->getCell('A'.$new_row)->setValue($df['no']);
                                $worksheet->getCell('B'.$new_row)->setValue($df["name"]);
                                $worksheet->getCell('C'.$new_row)->setValue($df['column']);
                                $worksheet->getCell('E'.$new_row)->setValue($df['tbl_re']);
                                $worksheet->getCell('G'.$new_row)->setValue($df['tbl_re_col']);
                                $new_row++;
                            }
                        }
                        break;
                    }
                }
            }
        } catch (Exception $e) {
            dump($table);
            dump($data_fa);
            dd($e->getMessage());
        }
    }
    // Sau khi export xong thì xóa sheet template ở file mẫu đi.
    $spreadsheet->setActiveSheetIndexByName('template');
    $sheetIndex = $spreadsheet->getActiveSheetIndex();
    $spreadsheet->removeSheetByIndex($sheetIndex);
    // Kết thúc xóa sheet template
    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
    $filename = 'db_structure_' . date('Ymd_His') . '.xlsx';
    $writer->save($filename);    

    dd('done');
});

Route::get('gen_db', function(){

    $sql_table_str = "
        SELECT
            TABLE_NAME,
            COLUMN_NAME,
            DATA_TYPE,
            (CASE  WHEN IS_NULLABLE = 'NO' THEN 'YES' ELSE '' END) as IS_NULLABLE,
            COLUMN_DEFAULT
            
        FROM
            INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = 'tostem_db_it_20210825'
        order by TABLE_NAME, ORDINAL_POSITION
    ";

    $data_tables = DB::connection(env('DB_CONNECTION_2', 'mysql2'))->select($sql_table_str);

    $sql_get_pri_unique = "
        select stat.table_schema as database_name,
               stat.table_name,
               stat.index_name,
               group_concat(stat.column_name
                    order by stat.seq_in_index separator ', ') as columns,
               tco.constraint_type
        from information_schema.statistics stat
        join information_schema.table_constraints tco
             on stat.table_schema = tco.table_schema
             and stat.table_name = tco.table_name
             and stat.index_name = tco.constraint_name
        where stat.table_schema = 'tostem_db_it_20210825'
        group by stat.table_schema,
                 stat.table_name,
                 stat.index_name,
                 tco.constraint_type
        order by stat.table_schema,
                 stat.table_name;
    ";

    $data_indexs = DB::connection(env('DB_CONNECTION_2', 'mysql2'))->select($sql_get_pri_unique);

    $tables = [];
    $datas = [];
    $no = 1;
    foreach ($data_tables as $dkey => $data_table) {
        $namejp = '';
        if ($data_table->COLUMN_NAME == 'updated_at' || $data_table->COLUMN_NAME == 'upd_datetime') {
            $namejp = '更新日時';
        } else if ($data_table->COLUMN_NAME == 'add_datetime' || $data_table->COLUMN_NAME == 'created_at') {
            $namejp = '登録日時';
        } else if ($data_table->COLUMN_NAME == 'upd_user_id') {
            $namejp = '更新者ＩＤ';
        } else if ($data_table->COLUMN_NAME == 'add_user_id') {
            $namejp = '登録者ＩＤ';
        } else {
            $namejp = ucwords(str_replace('_', " ", $data_table->COLUMN_NAME));
        }
        $table_name = strtolower($data_table->TABLE_NAME);
        $datas[$table_name][$dkey]['no'] = $dkey;
        $datas[$table_name][$dkey]['name_jp'] = $namejp;
        $datas[$table_name][$dkey]['column_name'] = $data_table->COLUMN_NAME;
        $datas[$table_name][$dkey]['data_type'] = $data_table->DATA_TYPE;
        $datas[$table_name][$dkey]['not_null'] = $data_table->IS_NULLABLE == 'YES' ? 'Yes':'';
        $datas[$table_name][$dkey]['column_default'] = $data_table->COLUMN_DEFAULT ?? '';
        $tables[$table_name] = $table_name;
    }

    $tbls = array_keys($tables);

    $str_q = '(\''.implode("','", $tbls).'\')';
    $sql_get_f = "
        SELECT 
          TABLE_NAME,COLUMN_NAME,CONSTRAINT_NAME, REFERENCED_TABLE_NAME,REFERENCED_COLUMN_NAME
        FROM
          INFORMATION_SCHEMA.KEY_COLUMN_USAGE
        WHERE
          REFERENCED_TABLE_SCHEMA = 'tostem_db_it_20210825' AND
          TABLE_NAME IN $str_q
        order by TABLE_NAME
    ";

    $data_fr = DB::connection(env('DB_CONNECTION_2', 'mysql2'))->select($sql_get_f);
    $data_frs = [];
    foreach ($data_fr as $key => $data_fr_value) {
        $table = $data_fr_value->TABLE_NAME;
        if (!isset($data_frs[$table])) {
            $no = 1;
        }
        $data_frs[$table][$key]['no'] = $no++;
        $data_frs[$table][$key]['name'] = $data_fr_value->CONSTRAINT_NAME;
        $data_frs[$table][$key]['column'] = $data_fr_value->COLUMN_NAME;
        $data_frs[$table][$key]['tbl_re'] = $data_fr_value->REFERENCED_TABLE_NAME;
        $data_frs[$table][$key]['tbl_re_col'] = $data_fr_value->REFERENCED_COLUMN_NAME;
    }

    foreach ($datas as $d_key => $data) {
        $no = 1;
        foreach ($data as $dkey => $dvalue) {
            $datas[$d_key][$dkey]['no'] = $no++;
        }
    }

    $datas_pri_uni = [];
    $current_table = '';
    foreach ($data_indexs as $kd => $data_index) {
        $tbl = $data_index->table_name;
        if ($data_index->constraint_type == 'PRIMARY KEY' || $data_index->constraint_type == 'UNIQUE') {
            $datas_pri_uni[$tbl][$kd]['no'] = $kd;
            $datas_pri_uni[$tbl][$kd]['name'] = $data_index->index_name;
            $datas_pri_uni[$tbl][$kd]['column'] = $data_index->columns;
            $datas_pri_uni[$tbl][$kd]['p_type'] = $data_index->constraint_type == 'PRIMARY KEY' ? 'Yes':'';
            $datas_pri_uni[$tbl][$kd]['u_type'] = $data_index->constraint_type == 'UNIQUE' ? 'Yes':'';
        }
    }

    foreach ($datas_pri_uni as $dp_k => $dp) {
        $no = 1;
        foreach ($dp as $dpkey => $value) {
            $datas_pri_uni[$dp_k][$dpkey]['no'] = $no++;
        }
    }


    //template file
    $file = storage_path('O2O_20210825.xlsx');
    $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($file);
    $sheetNames = $spreadsheet->getSheetNames();
    foreach ($sheetNames as $sheetName) {
        $table = $sheetName;
        if (!in_array($sheetName, $tables)) {
            if ($sheetName == 'history_import_product_selling_') {
                $table = 'history_import_product_selling_code_price';
            } else if ($sheetName == 't_data_convert_update_data_gies') {
                $table = 't_data_convert_update_data_giesta';
            } else {
                continue;
            }
        }

        $spreadsheet->setActiveSheetIndexByName($sheetName);
        $worksheet = $spreadsheet->getActiveSheet();

        $worksheet->getCell('C6')->setValue($table);
        $worksheet->getCell('F5')->setValue('MySQL');
        $data_struct = $datas[$table];
        $data_p = [];
        if (isset($datas_pri_uni[$table])) {
            $data_p = $datas_pri_uni[$table];
        }
        $data_fa = [];
        if (isset($data_frs[$table])) {
            $data_fa = $data_frs[$table];
        }

        foreach ($worksheet->getRowIterator() as $k_row => $row) {
            foreach( $row->getCellIterator() as $k_col => $cell ){
                $value = $cell->getCalculatedValue();
                if ($value == 'カラム情報') {
                    $new_row = $k_row + 2;
                    foreach ($data_struct as $data_s) {
                        $worksheet->insertNewRowBefore($new_row, 1);
                        $spreadsheet->getActiveSheet()->getStyle('A'.$new_row)->getFill()
                            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            ->getStartColor()->setARGB('ffffff');
                        $spreadsheet->getActiveSheet()->getStyle('B'.$new_row)->getFill()
                            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            ->getStartColor()->setARGB('ffffff');
                        $spreadsheet->getActiveSheet()->getStyle('C'.$new_row)->getFill()
                            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            ->getStartColor()->setARGB('ffffff');
                        $spreadsheet->getActiveSheet()->getStyle('D'.$new_row)->getFill()
                            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            ->getStartColor()->setARGB('ffffff');
                        $spreadsheet->getActiveSheet()->getStyle('E'.$new_row)->getFill()
                            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            ->getStartColor()->setARGB('ffffff');
                        $spreadsheet->getActiveSheet()->getStyle('F'.$new_row)->getFill()
                            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            ->getStartColor()->setARGB('ffffff');
                        $spreadsheet->getActiveSheet()->getStyle('G'.$new_row)->getFill()
                            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                            ->getStartColor()->setARGB('ffffff');
                        $worksheet->getCell('A'.$new_row)->setValue($data_s['no']);
                        $worksheet->getCell('B'.$new_row)->setValue($data_s['name_jp']);
                        $worksheet->getCell('C'.$new_row)->setValue($data_s['column_name']);
                        $worksheet->getCell('D'.$new_row)->setValue($data_s['data_type']);
                        $worksheet->getCell('E'.$new_row)->setValue($data_s['not_null']);
                        $worksheet->getCell('F'.$new_row)->setValue($data_s['column_default']);
                        $new_row++;
                    }
                }
                break;
            }
        }

        if (count($data_p)) {
            foreach ($worksheet->getRowIterator() as $k_row => $row) {
                foreach( $row->getCellIterator() as $k_col => $cell ){
                    $value = $cell->getCalculatedValue();
                    if ($value == 'インデックス情報') {
                        $new_row = $k_row + 2;
                        foreach ($data_p as $dp) {
                            $worksheet->insertNewRowBefore($new_row, 1);
                            $spreadsheet->getActiveSheet()->getStyle('A'.$new_row)->getFill()
                                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                ->getStartColor()->setARGB('ffffff');
                            $spreadsheet->getActiveSheet()->getStyle('B'.$new_row)->getFill()
                                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                ->getStartColor()->setARGB('ffffff');
                            $spreadsheet->getActiveSheet()->getStyle('C'.$new_row)->getFill()
                                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                ->getStartColor()->setARGB('ffffff');
                            $spreadsheet->getActiveSheet()->getStyle('D'.$new_row)->getFill()
                                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                ->getStartColor()->setARGB('ffffff');
                            $spreadsheet->getActiveSheet()->getStyle('E'.$new_row)->getFill()
                                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                ->getStartColor()->setARGB('ffffff');
                            $spreadsheet->getActiveSheet()->getStyle('F'.$new_row)->getFill()
                                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                ->getStartColor()->setARGB('ffffff');
                            $spreadsheet->getActiveSheet()->getStyle('G'.$new_row)->getFill()
                                ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                ->getStartColor()->setARGB('ffffff');
                            $worksheet->getCell('A'.$new_row)->setValue($dp['no']);
                            $worksheet->getCell('B'.$new_row)->setValue($dp['name']);
                            $worksheet->getCell('C'.$new_row)->setValue($dp['column']);
                            $worksheet->getCell('E'.$new_row)->setValue($dp['p_type']);
                            $worksheet->getCell('F'.$new_row)->setValue($dp['u_type']);
                            $new_row++;
                        }
                    }
                    break;
                }
            }
        }

        try {
            if (count($data_fa)) {
                foreach ($worksheet->getRowIterator() as $k_row => $row) {
                    foreach( $row->getCellIterator() as $k_col => $cell ){
                        $value = $cell->getCalculatedValue();
                        if ($value == '外部キー情報') {
                            $new_row = $k_row + 2;
                            foreach ($data_fa as $df) {
                                $worksheet->insertNewRowBefore($new_row, 1);
                                $spreadsheet->getActiveSheet()->getStyle('A'.$new_row)->getFill()
                                    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                    ->getStartColor()->setARGB('ffffff');
                                $spreadsheet->getActiveSheet()->getStyle('B'.$new_row)->getFill()
                                    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                    ->getStartColor()->setARGB('ffffff');
                                $spreadsheet->getActiveSheet()->getStyle('C'.$new_row)->getFill()
                                    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                    ->getStartColor()->setARGB('ffffff');
                                $spreadsheet->getActiveSheet()->getStyle('D'.$new_row)->getFill()
                                    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                    ->getStartColor()->setARGB('ffffff');
                                $spreadsheet->getActiveSheet()->getStyle('E'.$new_row)->getFill()
                                    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                    ->getStartColor()->setARGB('ffffff');
                                $spreadsheet->getActiveSheet()->getStyle('F'.$new_row)->getFill()
                                    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                    ->getStartColor()->setARGB('ffffff');
                                $spreadsheet->getActiveSheet()->getStyle('G'.$new_row)->getFill()
                                    ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
                                    ->getStartColor()->setARGB('ffffff');
                                $worksheet->getCell('A'.$new_row)->setValue($df['no']);
                                $worksheet->getCell('B'.$new_row)->setValue($df["name"]);
                                $worksheet->getCell('C'.$new_row)->setValue($df['column']);
                                $worksheet->getCell('E'.$new_row)->setValue($df['tbl_re']);
                                $worksheet->getCell('G'.$new_row)->setValue($df['tbl_re_col']);
                                $new_row++;
                            }
                        }
                        break;
                    }
                }
            }
        } catch (Exception $e) {
            dump($table);
            dump($data_fa);
            dd($e->getMessage());
        }
    }

    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save('write.xlsx');

    dd('done');
});
Route::get('/test', function () {
    return 'Router is working!';
});
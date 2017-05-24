//
//  ViewController.swift
//  TrainExcel
//
//  Created by 季 雲 on 2017/05/19.
//  Copyright © 2017 Ericji. All rights reserved.
//

import Cocoa

class ViewController: NSViewController {

    override func viewDidLoad() {
        super.viewDidLoad()
        
        let filePath = "/Users/y_ji/Desktop/Workbook1.xlsx"
        let odp = BRAOfficeDocumentPackage.open(filePath)
        let worksheet: BRAWorksheet = odp?.workbook.worksheets[0] as! BRAWorksheet
        let string = worksheet.cell(forCellReference: "A1").stringValue()
        
        

        // Do any additional setup after loading the view.
    }

    override var representedObject: Any? {
        didSet {
        // Update the view, if already loaded.
        }
    }


}


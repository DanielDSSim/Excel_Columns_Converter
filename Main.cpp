#include <iostream> 
#include <fstream>
#include <string>
#include "BasicExcel.hpp"

#define INPUTCOLS 5 //number of columns the desired format has 

std::string ideal_format[INPUTCOLS] = {"S/n", "Name", "Date", "Time", "Fun"}; //Desired headers

std::string new_Sheet_From_Template(std::string name) { //Makes a new sheet in the correct format ready to be filled in
	
	std::ifstream  src("Excel_Template.xls", std::ios::binary);
	name = "NEW" + name;
    std::ofstream  dst(name,   std::ios::binary);
	//std::cout << "Generating new spreadsheet\n";
    dst << src.rdbuf();
	
	return name;
}

/*void printoutstring(std::string str) {
	for(int i = 0; i < str.length(); i++) {
		std::cout << str[i] << endl;
	}
}*/

bool format_Check(YExcel::BasicExcel* Excel_File) { //Checks if the sheet is in the correct format by checking the column headers 1/true 0/false
	
	
	YExcel::BasicExcelWorksheet* Excel_Worksht =  Excel_File->GetWorksheet("Sheet1");
	
	if(Excel_Worksht->GetTotalCols() == INPUTCOLS) {
		std::string header;
		YExcel::BasicExcelCell* c;
		for(int i=0; i<INPUTCOLS; i++) {
			c = Excel_Worksht->Cell(0, i);
			header = c->GetString();
			if((ideal_format[i].compare(header)) != 0) {
				std::cout << header << " is being compared to " << ideal_format[i] << endl;
				std::cout << "Cell " << i+1 << " differs!" << endl;
				return false;
			}
		}
		return true;
	}
	//std::cout << "Worksheet format is different\n";
	return false;
}

void copyColumn(YExcel::BasicExcelWorksheet* new_sheet,YExcel::BasicExcelWorksheet* old_sheet,int posNew, int posOld) {
	
	//std::cout << old_sheet->GetTotalRows() << endl;
	for(int row=1; row < old_sheet->GetTotalRows(); row++) {
		if(old_sheet->Cell(row, posOld)->GetString() != 0) {
			new_sheet->Cell(row, posNew)->Set(old_sheet->Cell(row, posOld)->GetString());
			//std::cout << "copying row " << row << endl;
		}
		if(old_sheet->Cell(row, posOld)->GetInteger() != 0) {
			new_sheet->Cell(row, posNew)->Set(old_sheet->Cell(row, posOld)->GetInteger());
		}
		if(old_sheet->Cell(row, posOld)->GetDouble() != 0) {
			new_sheet->Cell(row, posNew)->Set(old_sheet->Cell(row, posOld)->GetDouble());
		} 
	}
}

bool convert_Excel_to_Template(std::string source_file) { //Converts the Excel file into the ideal_format 1/success 0/fail
	
	YExcel::BasicExcel eFile;
	
	bool check = eFile.Load(source_file.c_str());
	if(!check) {
		std::cout << "Failed to load " << source_file << endl;
		return 0;
	}
	
	YExcel::BasicExcel* file_ptr = &eFile;
	
	check = format_Check(file_ptr);
	if(check) {
		std::cout << "File is in correct format\n";
	} else {
		std::string NewName = new_Sheet_From_Template(source_file);
		YExcel::BasicExcelWorksheet* old_sheet = eFile.GetWorksheet("Sheet1"); //generates the sheet in question 
		std::string* old_Header_Contents = new string[old_sheet->GetTotalCols()]; //creates an array of size equal to col number in old sheet
		
		YExcel::BasicExcelCell* cell;
		for(int i=0; i < old_sheet->GetTotalCols(); i++) { //loop to fill the array with the headers for later searching and indexing 
			cell = old_sheet->Cell(0, i);
			old_Header_Contents[i] = cell->GetString();
			//std::cout << old_Header_Contents[i] << endl;
		}
		
		YExcel::BasicExcel NeFile;
		NeFile.Load(NewName.c_str());
		YExcel::BasicExcelWorksheet* new_sheet = NeFile.GetWorksheet("Sheet1");
		
		for(int i=0; i < INPUTCOLS; i++) {
			for(int j=0; j < old_sheet->GetTotalCols(); j++) {
				if(ideal_format[i] == old_Header_Contents[j]) {
					//std::cout << "Copying old sheet col " << j << " into new sheet col " << i << endl; 
					copyColumn(new_sheet, old_sheet, i, j);
				}
			}
		}
		
		check = NeFile.Save(); //save the translated file 
		if(!check) {
			std::cout << "SAVE ERROR\n";
		}
		
	}
	return 1;
}

int main() { //Take in user input for the file name and generates a spreadsheet with "NEW" prefix in the correct format
	
	std::string name;
	std::cout << "Input file name\n";
	std::cin >> name;
	bool check = convert_Excel_to_Template(name);
	
}
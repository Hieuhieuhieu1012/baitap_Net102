using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using ClosedXML;
using ClosedXML.Excel;




namespace baitapNet102
{
	public class Program
	{

		public static void Main(string[] args)
		{
			List<Class> clases = new List<Class>();
			Class classer = new Class();

			classer.Id = Guid.NewGuid();
			Console.WriteLine("Id: " + classer.Id);

			Console.Write("Nhập mã lớp : ");
			classer.MaLop = Console.ReadLine();

			Console.Write("Nhập tên lớp : ");
			classer.TenLop = Console.ReadLine();

			Console.Write("Số Lượng : ");
			classer.SoLuong = int.Parse(Console.ReadLine());

			clases.Add(classer);

			Console.Write("so luong mon hoc: ");
			int mon = Int32.Parse(Console.ReadLine());

			for (int i = 0; i < mon; i++)
			{
				
				MonHoc lop = new MonHoc();

				Console.Write("Nhập mã môn : ");
				lop.MaMonHoc = int.Parse(Console.ReadLine());

				Console.Write("Nhập tên môn : ");
				lop.TenMonHoc = Console.ReadLine();


				Console.Write("so luong Sinh Vien: ");
				int count = Int32.Parse(Console.ReadLine());
				List<Student> students = new List<Student>();
				for (int j = 0; j < count; j++)
				{
					Student employee = new Student();


					employee.Id = Guid.NewGuid();
					Console.WriteLine("Id: " +employee.Id);

					Console.Write("Ma Sinh Vien: ");
					employee.MaSV = Console.ReadLine();

					Console.Write("Tên Sinh Vien: ");
					employee.Name = Console.ReadLine();

					Console.Write("Ma Mon Hoc ");
					employee.MaMonHoc =  Int32.Parse( Console.ReadLine());

					Console.Write("Diem : ");
					employee.Diem = Int32.Parse(Console.ReadLine());

					Console.Write("So Buoi Nghi : ");
					employee.SoBuoiNghi = Int32.Parse(Console.ReadLine());

					Console.Write("Lop: ");
					employee.MaLop = Console.ReadLine();

					students.Add(employee);
					

				}

				Console.WriteLine("Thông tin các môn và sinh viên:");
				
					Console.WriteLine($"Môn: {lop.TenMonHoc} (Mã Môn: {lop.MaMonHoc})");
					var sinhVienTrongLop = students.Where(sv => sv.MaMonHoc == lop.MaMonHoc);
					foreach (var sv in sinhVienTrongLop)
					{
						Console.WriteLine($"  Mã SV: {sv.MaSV}, Tên SV: {sv.Name}, Lóp: {sv.MaLop}");
					}
				Console.WriteLine("Bạn có muốn xuất file Excel ko? 1- yes ; 2- no");
				int item = Int32.Parse(Console.ReadLine());
				if (item == 1)
				{

					string fileName = "DanhSachSinhVien.xlsx";
					string customPath = @"C:\Users\admin\Source\Repos\baitapNet102";
					string filePath = Path.Combine(customPath, fileName);
					ExportToExcel(students, filePath);

					Console.WriteLine("Xuất dữ liệu thành công.");

				}
				else
				{
					Console.WriteLine();
				}
			}

			Console.WriteLine("Nhap tên lớp cần tìm");
			string search = Console.ReadLine();
			var searchByName = clases.Where(c => c.TenLop == search);
		}

		static void ExportToExcel(List<Student> danhSachSinhVien, string filePath)
		{
			using (var workbook = new XLWorkbook())
			{
				var worksheet = workbook.Worksheets.Add("DanhSachSinhVien");

				
				worksheet.Cell(1, 1).Value = "Mã SV";
				worksheet.Cell(1, 2).Value = "Tên SV";
				worksheet.Cell(1, 3).Value = "Điểm";
				worksheet.Cell(1, 4).Value = "Mã Môn";
				worksheet.Cell(1, 4).Value = "Mã Lớp";


				for (int i = 0; i < danhSachSinhVien.Count; i++)
				{
					var sv = danhSachSinhVien[i];
					worksheet.Cell(i + 2, 1).Value = sv.MaSV;
					worksheet.Cell(i + 2, 2).Value = sv.Name;
					worksheet.Cell(i + 2, 3).Value = sv.Diem;
					worksheet.Cell(i + 2, 4).Value = sv.MaMonHoc;
					worksheet.Cell(i + 2, 4).Value = sv.MaLop;
				}

				
				worksheet.Columns().AdjustToContents();

				
				workbook.SaveAs(filePath);
			}
		}
	}
}

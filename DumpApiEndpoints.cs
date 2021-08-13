using SwiftExcel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Web;
using System.Web.Http;


namespace InternalRecursion.EndPointDump
{
	public class DumpApiEndpoints
	{
		public static DataTable DumpFromAssemblyToDataTable(Assembly Asm)
		{
			var OutputTable = new DataTable();

			OutputTable.Columns.Add("Controller", typeof(string));
			OutputTable.Columns.Add("Method", typeof(string));
			OutputTable.Columns.Add("Route", typeof(string));
			OutputTable.Columns.Add("Used", typeof(string));

			var RoutesByController = new Dictionary<Type, MethodInfo[]>();
			var RouteType = typeof(RouteAttribute);

			foreach (var ControllerType in Asm.GetTypes()
				.Where(Type => typeof(ApiController).IsAssignableFrom(Type)))
			{
				RoutesByController.Add(
					ControllerType,
					ControllerType.GetMethods().Where(Method => Method.IsPublic && Method.IsDefined(RouteType)).ToArray());
			}


			foreach (var Entry in RoutesByController)
			{

				foreach (var Method in Entry.Value)
				{
					var RouteAttribute = Method.GetCustomAttribute(RouteType) as RouteAttribute;

					OutputTable.Rows.Add(Entry.Key.Name, Method.Name, RouteAttribute.Template, "Unknown");

				}

			}


			return OutputTable;
		}


		public static void DumpFromAssemblyToExcel(Assembly Asm, string OutputPath)
        {
			var Dump = DumpFromAssemblyToDataTable(Asm);


			using (var Writer = new ExcelWriter(OutputPath))
			{
				Writer.Write("Controller", 1, 1);
				Writer.Write("Method", 2, 1);
				Writer.Write("Route", 3, 1);
				Writer.Write("Used", 4, 1);

				for (int j = 0; j < Dump.Rows.Count; ++j)
				{
					for (int k = 0; k < Dump.Columns.Count; ++k)
					{
						Writer.Write(Dump.Rows[j].ItemArray[k].ToString(), k + 1, j + 2);
					}

				}



			}

		}
	}
}
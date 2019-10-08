require 'win32ole'



class ExcelHandler


    attr_accessor :sh

    def initialize
        @app = WIN32OLE.new('Excel.Application')
        @app.Visible = true
        @wrk = @app.Workbooks.Add
        @sh = @wrk.Sheets.Add
    end

    def zoom
        @app.ActiveWindow.Zoom = 80
    end


    def put_matrix_into_excel matrix, start, mojeLambda = nil

        matrix.each_with_index do | l, r |
            l.each_with_index do | i, c |

                @sh.Cells( r + start[0], c + start[1]).Value = i
                mojeLambda.call( @sh.Cells( r + start[0], c + start[1]) ) if !mojeLambda.nil?
            end
        end
    end

    def filter_by_cell r, c
        @sh.Cells(r,c).AutoFilter
    end

    def autofit_by r, ile

        

        range_to_autofit = @sh.Range( @sh.Cells( r, 2 ), @sh.Cells( r, 2 + ile ) )
        range_to_autofit.Columns.Autofit

    end


end

class AdoDbHalnder
    def initialize 
        @cnn = WIN32OLE.new('ADODB.Connection')

        @cfg = []
        @cfg << "Provider=Microsoft.ACE.OLEDB.12.0;" 
        @cfg << "Data Source=X:\\PLGLI-3-Exchange\\SoE\\Cost\\CONTEVAN\\CONTEVAN.accdb;"
        @cfg << "User Id=admin;Password="


        @adSchemaColumns = 4
        @adSchemaTables = 20
    end


    def change_path_with_double_slash new_path
        @cfg[1] = "Data Source=#{new_path}\\CONTEVAN.accdb;"
    end

    def open
        @cnn.Open( @cfg[0].to_s + @cfg[1].to_s )
       
    end


    def list_of_tables
        
        @rst = @cnn.OpenSchema(@adSchemaTables)

        data = @rst.GetRows.transpose
        @rst.Close
        data
    end


    def make_query str
        @rst = WIN32OLE.new("ADODB.Recordset")
        @rst.Open(str, @cnn)
            # @rst.Open("SELECT DISTINCT PLT FROM [LA KW 37 19];", @cnn)

        data = @rst.GetRows.transpose
        @rst.Close
        data
    end
end
require './win-32-ole.rb'
require './date-handler.rb'
require './mgo-handler.rb'


class ContevanController

    attr_accessor :input_data, :cws, :parsed_cws, :re_parsed

    def initialize
        @input_data = []
        @cws = []
        @parsed_cws = []
        @re_parsed = []

        @la_labels = [ 'Duns','PN','PLT','Supplier','PName','FabAuth','MatAuth','Fup','IssueDate','CumProd','PCT','CumShip','CumReq' ]
        @nb_labels = [ 'Fup','PN','PLT','PName','DOCKCD','IssueDate','INVSTOCK','YTDREC','TOTADJ','PRECUM', 'BEG_BAL', 'DOH', 'BANK' ]


        
    end


    def parse_cws

        @parsed_cws = []
        @re_parsed = []

        
        t1 = @cws[0].to_s
        t2 = @cws[1].to_s


        # just to be sure
        if t1.to_i <= t2.to_i then

            dh = DateHandler.new
            monday_from_t1 = dh.get_monday_from_ t1
            monday_from_t2 = dh.get_monday_from_ t2

            #monday_from_t1 = monday_from_t1.strftime("%F")
            #monday_from_t2 = monday_from_t2.strftime("%F")


            #p [ monday_from_t1, monday_from_t2 ]
            #STDIN.gets

            until monday_from_t1 > monday_from_t2

                monday_from_t1_year = monday_from_t1.year
                cw = monday_from_t1.cwxx
                mm = monday_from_t1.month


                tmp_ycw = "#{monday_from_t1_year}00".to_i
                tmp_ycw += monday_from_t1.cwxx.to_i
                @parsed_cws << tmp_ycw
                monday_from_t1 += ( 7 * Time.jeden_dzien )
            end
        else
            @parsed_cws = []
        end


        if @parsed_cws.size > 0

            @parsed_cws.each do |e|
                
                #@re_parsed.insert(0,"KW " + e.to_s[4...6] + " " + e.to_s[2...4])
                @re_parsed << "KW " + e.to_s[4...6] + " " + e.to_s[2...4]
            end
        end


        #p @re_parsed
        #STDIN.gets

    end


    def run_logic type_of_rep, input_list, m_cws, is_cum


        #p input_list
        #p @input_data
        #p m_cws
        #p @cws


        inner_sh = nil


        # good simple list checking by == , which is great
        if input_list == @input_data and m_cws == @cws then

            #puts "OK!" # remember that in shoes you can not use puts
            parse_cws


            if @re_parsed.size > 0 then

                inner_sh = run_rqms(input_list, m_cws, is_cum)  if type_of_rep == "RQM"
                inner_sh = run_sched(input_list, m_cws, is_cum) if type_of_rep == "SCHED"
            end
        else
            # sth went really wrong
        end


        inner_sh
    end

    def some_formatting my_excel, typ

        my_excel.filter_by_cell 4, 2
        #my_excel.autofit_by 4, (40 + @re_parsed.size)
        my_excel.zoom

        f_arr = []
        if typ == "LA"
            f_arr = [ ["B:B",11], ["C:C",9], ["D:D",10], ["E:E",6], ["F:F",25], ["G:G",15], ["H:H", 11], ["I:I", 11], ["J:J",5], ["K:K", 12], ["L:L", 12] ]
            f_arr << ["N:N", 11]
            f_arr << ["O:O", 10]

        elsif typ == "NB"
            f_arr = [ ["B:B",11], ["C:C",6], ["D:D",10], ["E:E",6], ["F:F",25], ["G:G", 13], ["H:H", 13], ["I:I", 13], ["J:J", 13],  ["K:K", 13], ["L:L", 12], ["M:M", 11] ]
        end

        f_arr.each { |x| my_excel.sh.Columns(x[0]).ColumnWidth  = x[1] }

        # from P now same till end
        #For VerticalAlignment:
        #Top:    -4160
        #Center: -4108
        #Bottom: -4107
        #And HorizontalAlignment:
        #Left:    -4131
        #Center:  -4108
        #Right:   -4152
        xlLeft = -4131

        l_cw = my_excel.sh.Cells(4, 16)
        while !l_cw.Value.nil? && l_cw.Value.strip != ""
            l_cw.HorizontalAlignment = xlLeft
            l_cw.WrapText = false
            l_cw.EntireColumn.ColumnWidth = 14

            l_cw = l_cw.Offset(0,1)

        end 

    end



    def prepare_input_for_mgo_handler type_str

        if !@my_excel.nil? then

            # plt = @my_excel.sh.Cells(5,5).Value
        end


        arr = []
        arr
    end


    def run_rqms input_list, m_cws, is_cum
        
        
        output = []
        
        app = AdoDbHalnder.new
        app.open

        prefix = "NB"

        #p @input_data
        #p @re_parsed

        @re_parsed.each_with_index do | str_cw, i1|

            access_tbnm = "#{prefix} #{str_cw}"

            @input_data.each_with_index do |e, i2|
                
                # example:
                # "SELECT * FROM [LA KW 37 19] WHERE (([Part No] = 39203558) AND ([PLT] = 'PO' OR [PLT] = 'EP'));"
                plt = e[0].strip
                pn = e[1].strip
                q = ""
                if (not pn.empty?) and (not plt.empty?) then
                    q = "SELECT DISTINCT * FROM [#{access_tbnm}] WHERE (([Part No] = #{e[1]}) AND ([PLT] = '#{e[0]}'));"
                elsif (not pn.empty?) and (plt.empty?) then
                    q = "SELECT DISTINCT * FROM [#{access_tbnm}] WHERE (([Part No] = #{e[1]}));"
                else
                    q = ""
                end


                if (not q.empty?) then


                    double_arr = app.make_query(q)

                    double_arr.each do |rqm_line|
                        
                        rqm_line_bank = rqm_line.pop
                        rqm_line_doh = rqm_line.pop
                        rqm_line_beg_bal = rqm_line.pop

                        rqm_line.insert( @nb_labels.index('BEG_BAL') , rqm_line_beg_bal ) 
                        rqm_line.insert( @nb_labels.index('DOH') ,  rqm_line_doh ) 
                        rqm_line.insert( @nb_labels.index('BANK') ,  rqm_line_bank ) 
                    end

                    double_arr.each do |i_inside_arr| 
                        i_inside_arr.insert(0, str_cw)
                        i1.times { |iterator| i_inside_arr.insert(@nb_labels.size + 1, " ") }
                    end

                    output << double_arr

                end
            end
        end





        my_excel = ExcelHandler.new
        dh = DateHandler.new
        my_lambda = lambda do |cell| 
            cell.Font.Bold = true
            cell.Font.Color = 0xFFFFFF
            cell.Interior.Color = 0x000000;
        end

        x = 4
        raw_la_labels = @nb_labels.insert(0, "table")
        (40 + @re_parsed.size - 1).times do |i|
             t = @cws[0].to_s
             mon = dh.get_monday_from_(t)
             mon += ( i * 7 * Time.jeden_dzien )
             
             raw_la_labels << "#{mon.year}_CW#{mon.cwxx}"
        end
        boxed_la_labels = [[raw_la_labels]]
        boxed_la_labels.each_with_index do |data, i |
            my_excel.put_matrix_into_excel data, [ x, 2 ], my_lambda
        end

        x = 5
        output.each_with_index do | data, i|

            my_excel.put_matrix_into_excel data, [ x, 2 ]
            x = x + data.length
            
        end

        some_formatting my_excel, prefix

        return my_excel.sh;

    end

    def run_sched input_list, m_cws, is_cum

        # $stderr.puts "run_sched -> test"
        #input_list here same as @input_data  -> [["EP", "39203558", "139956"], ["PO", "39203558", "139956"], ["EP", "39203559", "139956"]]
        # m_cws same as @cws -> ["201936", "201937"]

        # here will store all the data
        output = []
        
        app = AdoDbHalnder.new
        app.open

        prefix = "LA"

        #p @input_data
        #p @re_parsed

        @re_parsed.each_with_index do | str_cw, i1|

            access_tbnm = "#{prefix} #{str_cw}"

            @input_data.each_with_index do |e, i2|
                
                # example:
                # "SELECT * FROM [LA KW 37 19] WHERE (([Part No] = 39203558) AND ([PLT] = 'PO' OR [PLT] = 'EP'));"
                plt = e[0].strip
                pn = e[1].strip
                duns = e[2].strip
                q = ""
                if (not duns.empty?) and (not pn.empty?) and (not plt.empty?) then
                    q = "SELECT DISTINCT * FROM [#{access_tbnm}] WHERE (([Duns] = #{e[2]}) AND ([Part No] = #{e[1]}) AND ([PLT] = '#{e[0]}'));"
                elsif (duns.empty?) and (not pn.empty?) and (not plt.empty?) then
                    q = "SELECT DISTINCT * FROM [#{access_tbnm}] WHERE (([Part No] = #{e[1]}) AND ([PLT] = '#{e[0]}'));"
                elsif (duns.empty?) and (not pn.empty?) and (plt.empty?) then
                    q = "SELECT DISTINCT * FROM [#{access_tbnm}] WHERE (([Part No] = #{e[1]}));"
                elsif (not duns.empty?) and (pn.empty?) and (plt.empty?) then
                    q = "SELECT DISTINCT * FROM [#{access_tbnm}] WHERE (([Duns] = #{e[2]}));"   
                else
                    q = ""
                end


                if (not q.empty?) then


                    double_arr = app.make_query(q)
                    double_arr.each do |i_inside_arr| 
                        i_inside_arr.insert(0, str_cw)
                        i1.times { |iterator| i_inside_arr.insert(@la_labels.size + 1, " ") }
                    end

                    output << double_arr

                end
            end
        end





        my_excel = ExcelHandler.new
        dh = DateHandler.new
        my_lambda = lambda do |cell| 
            cell.Font.Bold = true
            cell.Font.Color = 0xFFFFFF
            cell.Interior.Color = 0x000000;
        end

        x = 4
        raw_la_labels = @la_labels.insert(0, "table")
        (40 + @re_parsed.size - 1).times do |i|
             t = @cws[0].to_s
             mon = dh.get_monday_from_(t)
             mon += ( i * 7 * Time.jeden_dzien )
             
             raw_la_labels << "#{mon.year}_CW#{mon.cwxx}"
        end
        boxed_la_labels = [[raw_la_labels]]
        boxed_la_labels.each_with_index do |data, i |
            my_excel.put_matrix_into_excel data, [ x, 2 ], my_lambda
        end

        x = 5
        output.each_with_index do | data, i|

            my_excel.put_matrix_into_excel data, [ x, 2 ]
            x = x + data.length
            
        end

        some_formatting my_excel, prefix


        return my_excel.sh;
    end



end


def test_cw_parse
    cc = ContevanController.new
    cc.input_data = [
        ["1","2","3"],
        ["4","5","6"]
    ]

    cc.cws = ["201940", "201940"]
    cc.parse_cws

    # p cc.re_parsed
end

#test_cw_parse


def test_first_query_schedule

    cc = ContevanController.new
    cc.input_data = [
        ["EP","39203558","139956"],
        ["PO","39203558","139956"],
        ["EP","39203559","139956"]
    ]

    cc.cws = ["201936", "201937"]
    cc.parse_cws


    #run_logic type_of_rep, input_list, m_cws, is_cum, asns
    output = cc.run_logic( "SCHED", cc.input_data, cc.cws, false)


    #p output
    # output.each { |item|  p item; p "\n\n\n";  }

    my_excel = ExcelHandler.new
    output.each_with_index do | data, i|
        #p [ 2 , 2 + i * 3 ]
        my_excel.put_matrix_into_excel data, [ 2 + i, 2 ]
    end
end


# test_first_query_schedule




# test main logic for main button
def test_main_logic
    
    m_input_data = [["","","139956"]]
    
    m_cws = ["201938", "201940"]
    typeString = "SCHED"

    c_ctrl = ContevanController.new
    c_ctrl.input_data = m_input_data
    c_ctrl.cws = m_cws
    #c_ctrl.run_logic "RQM", c_ctrl.input_data, c_ctrl.cws, false, false
    sh = c_ctrl.run_logic typeString, c_ctrl.input_data, c_ctrl.cws, false

    mgo_handler = MgoHandler.new(sh)
    input_for_mgo_handler = mgo_handler.prepare_input_for_mgo_handler()

end


#test_main_logic



def test_on_rqm
    c_ctrl = ContevanController.new
    c_ctrl.input_data = [["","39203558",""]]
    #c_ctrl.input_data = [["","39203558",""]]
    #c_ctrl.input_data = [["PO","39203558",""]]
    #c_ctrl.input_data = [["","","139956"]]
    c_ctrl.cws = ["201930", "201937"]
    c_ctrl.run_logic "RQM", c_ctrl.input_data, c_ctrl.cws, false
    c_ctrl.run_logic "SCHED", c_ctrl.input_data, c_ctrl.cws, false
end

#test_on_rqm



def run_rqm m_input_data, m_cws, typeString, cum, asn
    c_ctrl = ContevanController.new
    c_ctrl.input_data = m_input_data
    c_ctrl.cws = m_cws
    sh = c_ctrl.run_logic typeString, c_ctrl.input_data, c_ctrl.cws, cum

    if asn == true then

        # in this constructor we already have prepare_input_for_mgo_handler
        # also: output_arr = []
        mgo_handler = MgoHandler.new(sh)
        ms = MS9PO400.new mgo_handler.sess0
        ms.open
        mgo_handler.input_arr.each do |pair|
            ms.set_plt_and_part_and_submit pair[0], pair[1]

            line = 1
            while (ms.get_qty(line) != 0) && (line < 9)
                mgo_handler.output_arr << OutputItem.new( pair[0], pair[1], ms.get_qty(line), ms.get_sid(line), ms.get_yyyycw_from_eda(line) )
    
                line += 1
            end
        end


        p mgo_handler.output_arr
    end
end

def run_sched m_input_data, m_cws, typeString, cum, asn
    c_ctrl = ContevanController.new
    c_ctrl.input_data = m_input_data
    c_ctrl.cws = m_cws
    #c_ctrl.run_logic "RQM", c_ctrl.input_data, c_ctrl.cws, false, false
    sh = c_ctrl.run_logic typeString, c_ctrl.input_data, c_ctrl.cws, cum

    if asn == true then

        # in this constructor we already have prepare_input_for_mgo_handler
        # also: output_arr = []
        mgo_handler = MgoHandler.new(sh)
        ms = MS9PO400.new mgo_handler.sess0
        ms.open
        mgo_handler.input_arr.each do |pair|
            ms.set_plt_and_part_and_submit pair[0], pair[1]

            line = 1
            while (ms.get_qty(line) != 0) && (line < 9)

                #puts "line #{line}" # test-case OK

                mgo_handler.output_arr << OutputItem.new( pair[0], pair[1], ms.get_qty(line), ms.get_sid(line), ms.get_yyyycw_from_eda(line) )
    
                line += 1

                # this statement working only if screen is full of asns
                if line == 9 then
                    line = ms.press_f8
                end
            end
        end


        p mgo_handler.output_arr
    end
end


def test_on_run_sched
    run_sched [["","13520813",""]], ["201940", "201940"], "SCHED", false, true
end

# test_on_run_sched
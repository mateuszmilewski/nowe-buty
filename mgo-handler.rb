require 'win32ole'
require './date-handler.rb'



class OutputItem

    attr_accessor :plt, :pn, :qty, :nm, :yyyycw, :mgo_screen_name, :d

    def initialize mplt, mpn, mqty, mnm, mycw, m_mgo_screen_name, md
        @plt = mplt
        @pn = mpn
        @qty = mqty
        @nm = mnm
        @yyyycw = mycw
        @mgo_screen_name = m_mgo_screen_name.to_s
        @d = md
    end

    def stringify
        return "#{@yyyycw} #{@nm} #{@plt} #{@pn} #{@qty} #{@src}"
    end
end


class MgoHandler

    attr_accessor :sess0, :sh, :input_arr, :output_arr

    def initialize ext_sh
        @app = WIN32OLE.new('EXTRA.System')
        @ss = @app.Sessions
        @sess0 = @app.ActiveSession
        @scr = @sess0.Screen


        

        @po400 = MS9PO400.new @sess0
        @ph100 = MS9PH100.new @sess0

        #
        #
        @sh = ext_sh
        @input_arr = prepare_input_for_mgo_handler()

        # p @input_arr

        @output_arr = []
    end


    def prepare_input_for_mgo_handler

        arr = []
        wiersz = 5

        while !@sh.Cells(wiersz , 5).Value.nil?

            pn = @sh.Cells(wiersz , 4).Value
            plt = @sh.Cells(wiersz , 5).Value
            arr << [plt, pn]
            wiersz += 1

        end
        
        
        arr.uniq!
        arr
    end
end



class MGO_SCREEN

    attr_accessor :s, :month_arr


    def initialize sesja
        @s = sesja

        @plt = nil
        @pn = nil

        @month_arr = ['JA', 'FE', 'MR', 'AP', 'MY', 'JN', 'JL', 'AU', 'SE', 'OC', 'NO', 'DE']
    end


    def raw_get(wiersz, kolumna, ile)

        to_return = ""

        raw_str = @s.Screen.GetString( wiersz, kolumna, ile)

        if raw_str.strip != "" then
            to_return = raw_str.strip.to_i
        end

        to_return
    end


    def set_plt_and_part_and_submit m_plt, m_part_number
        @plt = m_plt
        @pn = m_part_number
    end


    def open
        @s.Screen.SendKeys "<Clear>"
        wait_for_mgo
        sleep 0.1
        #
        @s.Screen.SendKeys "#{self.class} <Enter>"
        wait_for_mgo
        sleep 0.1
    end

    def class_name
        #puts "#{self.class}"

        self.class
    end


    def wait_for_mgo
        sleep 0.1
        @s.Screen.WaitHostQuiet(3000)
        sleep 0.1

    end
end


class MS9PO400 < MGO_SCREEN
    def initialize sesja
        super(sesja)


        @@only_one_screen = "I5487"
        @@no_shipments = "I6155"
        @@scan_sth = "I6293"
    end

    def set_plt_and_part_and_submit m_plt, m_part_number
        super( m_plt, m_part_number )

        #plt
        @s.Screen.PutString(@plt, 3, 7)
        #pn
        @s.Screen.PutString(@pn, 3, 19)
        #kanban
        @s.Screen.PutString("    ", 3, 35)

        @s.Screen.SendKeys "<Enter>"
        wait_for_mgo
    end


    def press_f8

        finally_what = 9

        if @@only_one_screen == get_screen_info() then
            finally_what = 9
        elsif @@no_shipments == get_screen_info() then
            finally_what = 9
        elsif @@scan_sth == get_screen_info() then
            finally_what = 9
        else

            @s.Screen.SendKeys "<pf8>"
            wait_for_mgo

            finally_what = 1
        end

        #puts "press_f8 #{finally_what?}"

        finally_what
    end

    def get_screen_info

        info = @s.Screen.GetString( 22, 2, 5)
    end


    def get_yyyycw_from_eda(line)

        cw_eda = nil

        raw_eda = @s.Screen.GetString( 6 + 2 * (line - 1), 46, 6)

        if raw_eda != "______" then

            dd = raw_eda[0,2].to_i
            mm = @month_arr.index(raw_eda[2,2]) + 1
            yyyy = 2000 + raw_eda[4,2].to_i

            t = Time.new(yyyy, mm, dd)
            cw_eda = "#{yyyy}#{t.cwxx}"
            
        end

        cw_eda
    end


    def get_date(line)

        d = ""
        raw_str = @s.Screen.GetString(6 + 2 * (line - 1), 14, 6)

        if raw_str.strip != "" then
            d = raw_str.strip.to_s
        end

        d
    end

    def get_delivery_date(line)

        d = ""
        raw_str = @s.Screen.GetString(6 + 2 * (line - 1), 46, 6)

        if raw_str.strip != "" then
            d = raw_str.strip.to_s
        end

        d
    end

    def get_yyyycw_from_sdate(line)

        cw_sdate = nil

        raw_sdate = @s.Screen.GetString( 6 + 2 * (line - 1), 14, 6)

        if raw_sdate.strip != "" then

            dd = raw_sdate[0,2].to_i
            mm = @month_arr.index(raw_sdate[2,2]) + 1
            yyyy = 2000 + raw_sdate[4,2].to_i

            t = Time.new(yyyy, mm, dd)
            cw_sdate = "#{yyyy}#{t.cwxx}"
            
        end

        cw_sdate
    end


    def get_yyyycw_from_delivery_date(line)

        cw_sdate = nil

        raw_sdate = @s.Screen.GetString( 6 + 2 * (line - 1), 46, 6)

        if raw_sdate.strip != "" then

            dd = raw_sdate[0,2].to_i
            mm = @month_arr.index(raw_sdate[2,2]) + 1
            yyyy = 2000 + raw_sdate[4,2].to_i

            t = Time.new(yyyy, mm, dd)
            cw_sdate = "#{yyyy}#{t.cwxx}"
            
        end

        cw_sdate
    end

    def get_qty(line)

        qty = 0

        raw_str = @s.Screen.GetString( 6 + 2 * (line - 1), 5, 8)

        if raw_str.strip != "" then
            qty = raw_str.strip.to_i
        end

        qty
    end

    def get_sid(line)

        sid = ""

        raw_str = @s.Screen.GetString( 6 + 2 * (line - 1), 60, 10)

        if raw_str.strip != "" then
            sid = raw_str.strip
        end

        sid
    end
end


class MS9PH100 < MGO_SCREEN

    attr_accessor :map_for_years

    def initialize sesja
        super(sesja)

        @@recv_string = "RECV"
        @date_for_init_config = Time.dzis - Time.jeden_dzien * 90
        @map_for_years = make_a_map_for_those_90_days()
        @date_for_init_config_parsed = "#{@date_for_init_config.day}.#{@month_arr[@date_for_init_config.month - 1]}.#{@date_for_init_config.year.to_i - 2000}"

        #p "test date already with mgo format: #{@date_for_init_config.year.to_i - 2000}.#{@month_arr[@date_for_init_config.month - 1]}.#{@date_for_init_config.day}"

        # R6086: INQUIRY COMPLETE 
        @@inquiry_complete = "R6086"
        # R6021: MORE DATA TO DISPLAY
        @@more_data_to_display = "R6021"
        
    end

    def make_a_map_for_those_90_days()
        # this function returns array with pairs year, month for screen data where we only have pattern day-month (26SE for example)
        # it means we do not know directly if we are in the right year for proper recv

        mapa = []

        
        # y1
        year_for_first_checked_day = (Time.dzis - Time.jeden_dzien * 90).year

        # y2
        current_year = Time.dzis.year


        m1 = (Time.dzis - Time.jeden_dzien * 90).month
        m2 = Time.dzis.month

        # this is simple scenario
        if current_year == year_for_first_checked_day then

            # mapa = []
            y = year_for_first_checked_day.to_i
            m1.upto(m2) do |month|
                if month == 1 then
                    y += 1
                end
                mapa << [month, @month_arr[month - 1], y]
            end


        # the only possibility is when current year is + 1
        elsif current_year > year_for_first_checked_day then

            # scenario, when go back to the prev year
            #
            y = year_for_first_checked_day.to_i
            m1.upto(m2) do |month|
                if month == 1 then
                    y += 1
                end
                mapa << [month, @month_arr[month - 1], y]
            end

            #
            #
        else
            # not possible
            mapa = []
        end

        mapa

    end


    def press_f8

        finally_what = 10

        if @@inquiry_complete == get_screen_info() then
            finally_what = 10
        elsif @@more_data_to_display == get_screen_info() then

            @s.Screen.SendKeys "<pf8>"
            wait_for_mgo

            finally_what = 2
        else

            @s.Screen.SendKeys "<pf8>"
            wait_for_mgo

            finally_what = 2
        end

        #puts "press_f8 #{finally_what?}"

        finally_what
    end

    def get_screen_info

        info = @s.Screen.GetString( 22, 2, 5)
    end


    def set_plt_and_part_and_submit m_plt, m_part_number
        super( m_plt, m_part_number )

        # puts "#{@plt}  #{@pn}"
        # gets
        wait_for_mgo
        @s.Screen.PutString(@plt, 4, 6)
        @s.Screen.PutString(@pn, 4, 17)
        @s.Screen.PutString("        ", 6, 8)
        @s.Screen.PutString(@date_for_init_config_parsed, 6, 8)
        # @s.Screen.PutString("    ", 7, 8)
        @s.Screen.PutString(@@recv_string, 7, 8)

        @s.Screen.SendKeys "<Enter>"
        wait_for_mgo
    end

    def move_right_for_shp_dt

        loop do
            shp = @s.Screen.GetString(9, 70, 3)
            break if shp == "SHP"
            # F11 move to right
            @s.Screen.SendKeys "<pf11>"
            wait_for_mgo
        end

    end

    def get_sdate(line)

        d = ""
        raw_str = @s.Screen.GetString(11 + 1 * (line - 1), 70, 8)

        if raw_str.strip != "" then
            d = raw_str.strip.to_s
        end

        d
    end

    def get_date(line)

        d = ""
        raw_str = @s.Screen.GetString(11 + 1 * (line - 1), 30, 4)

        if raw_str.strip != "" then
            d = raw_str.strip.to_s
        end

        d
    end

    def get_yyyycw_from_ms9ph100_sdate(line)
        cw_sdate = nil

        d = ""
        raw_str = @s.Screen.GetString(11 + 1 * (line - 1), 70, 8)

        if raw_str.strip != "" then
            dd = raw_str[0,2].to_i
            mm = @month_arr.index(raw_str[2,2]) + 1
            yyyy = raw_str[4,4]

            t = Time.new(yyyy, mm, dd)
            cw_sdate = "#{yyyy}#{t.cwxx}"
        end


        cw_sdate
        
    end


    def get_yyyycw_from_ms9ph100_date(line)

        cw_sdate = nil

        raw_sdate = @s.Screen.GetString( 11 + 1 * (line - 1), 30, 4)

        if raw_sdate.strip != "" then

            dd = raw_sdate[0,2].to_i
            mm = @month_arr.index(raw_sdate[2,2]) + 1

            if @map_for_years.size > 0 then
                yyyy = (@map_for_years.select { |inner_arr| inner_arr[1].to_s == raw_sdate[2,2].to_s }).first[2]
            end

            t = Time.new(yyyy, mm, dd)
            cw_sdate = "#{yyyy}#{t.cwxx}"
            
        end

        cw_sdate
    end



    def get_qty(line)

        qty = 0

        raw_str = @s.Screen.GetString( 11 + 1 * (line - 1), 35, 10)

        if raw_str.strip != "" then
            raw_str.gsub! ",", ""
            raw_str.gsub! ".", ""
            qty = raw_str.strip.to_i

        end

        qty
    end

    def get_sid(line)

        sid = ""

        raw_str = @s.Screen.GetString( 11 + 1 * (line - 1), 55, 10)

        if raw_str.strip != "" then
            sid = raw_str.strip
        end

        sid
    end
end



def test_mgo
    m = MgoHandler.new nil
    ms = MS9PO400.new m.sess0
    #ms.class_name # OK
    ms.open
    input_arr = [["PO", "39203559"], ["PO", "39203540"], ["PO", "39203558"]]

    output_arr = []

    input_arr.each do |pair|
        ms.set_plt_and_part_and_submit pair[0], pair[1]

        line = 1
        while (ms.get_qty(line) != 0) && (line < 9)
            output_arr << OutputItem.new( pair[0], pair[1], ms.get_qty(line), ms.get_sid(line), ms.get_yyyycw_from_eda(line) )

            line += 1
        end
    end

    output_arr.each { |x| puts x.stringify }

    
end


#test_mgo



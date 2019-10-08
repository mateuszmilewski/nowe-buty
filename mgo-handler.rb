require 'win32ole'
require './date-handler.rb'



class OutputItem

    attr_accessor :plt, :pn, :qty, :nm, :yyyycw

    def initialize mplt, mpn, mqty, mnm, mycw
        @plt = mplt
        @pn = mpn
        @qty = mqty
        @nm = mnm
        @yyyycw = mycw
    end

    def stringify
        return "#{@yyyycw} #{@nm} #{@plt} #{@pn} #{@qty}"
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

    attr_accessor :s


    def initialize sesja
        @s = sesja

        @plt = nil
        @pn = nil

        @month_arr = ['JA', 'FE', 'MR', 'AP', 'MY', 'JN', 'JL', 'AU', 'SE', 'OC', 'NO', 'DE']
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
        puts "#{self.class}"
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
    def initialize sesja
        super(sesja)
    end

    def set_plt_and_part_and_submit m_plt, m_part_number
        super( m_plt, m_part_number )


        # puts "#{@plt}  #{@pn}"
        # gets
        wait_for_mgo
        @s.Screen.PutString(@plt, 4, 6)
        @s.Screen.PutString(@pn, 4, 17)

        @s.Screen.SendKeys "<Enter>"
        wait_for_mgo
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



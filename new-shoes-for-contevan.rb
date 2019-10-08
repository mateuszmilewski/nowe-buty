require './date-handler.rb'
require './contevan-controller.rb'

dh = DateHandler.new


Shoes.app(title: "Klapaucius!;!;!;!;!;!;!;!;!;!;", width: 600, height: 620, resizable: false ) {


    # Shoes::show_console
    # $stderr.puts "hello" # working

    background white
    fill black

    stack {
        background white
        fill black
        subtitle "New shoes for Contevan!"
        @head_img = image "foxy-2.png"
        #@head_img = image "http://poignant.guide/images/the.foxes-2.png"

        @temp_txt = ""

        flow {
            para " Plt:    "
            @edit_plt = edit_line :width => 40
            para "    PN:  "
            @edit_pn = edit_line :width => 80
            para "    Duns:  "
            @edit_duns = edit_line :width => 100
            para "     "
            @add = button "Add"
        }

        flow {
            para " Lista: "
            @lista = edit_box :height => 80 , :width => 343

            @check_input_list = button "Sprawdź", :height => 80, :width => 80, :align => "center"
            @clear_list = button "Wyczyść", :height => 80, :width => 80, :align => "center"

            @check_input_list.click {

                checked = "Sprawdzone!"

                tmp_arr_for_validation = @lista.text.split(/\n/)

                if tmp_arr_for_validation.size == 0 then
                    checked = checked + "; " + "lista jest pusta!"
                end


                #usun zera z przodu dla PN and DUNS

                new_tmp_arr_for_validation = []
                if tmp_arr_for_validation.size > 0 then

                    checked = checked + "; " + "possible removal of leading zeros"

                    del_leading_zeros = -> {
                        tmp_arr_for_validation.each { |line|

                            array_for_this_line = line.split(";")
                            array_for_this_line.each { |element_in_line|
                                element_in_line.sub!(/^0*/, '')
                            }
                            new_tmp_arr_for_validation << array_for_this_line.join(";")
                        }
                    }
                    del_leading_zeros.call
                end

                #$stderr.puts tmp_arr_for_validation
                tmp_arr_for_validation = new_tmp_arr_for_validation



                primary_validation_with_semicolon = true
                tmp_arr_for_validation.each do |line|
                    
                    if line.split(";").size == 3 
                        #OK nothing to do
                        #elsif( line.split(";").size = 3 && line[-1] != ";" )
                    elsif line.split(";").size == 2

                        if line[-1] == ";" then
                            # OK
                        else
                            primary_validation_with_semicolon = false
                        end

                    else
                        # at least one line is in wrong format
                        primary_validation_with_semicolon = false
                    end
                end

                if !primary_validation_with_semicolon then
                    checked = checked + "; złe formatowanie listy wejściowe MUSZĄ BYĆ DWA ŚREDNIKI! -> wzorzec dla każdej linii: PLT;PN;DUNS -> wszystkie opcjonalne + DUNS nie jest dostępny dla RQMsow"
                end

                @log_text.replace "#{checked}, lista: #{tmp_arr_for_validation}"
                @moj_text.replace "#{checked}, lista: #{tmp_arr_for_validation}"
            }

            @clear_list.click {
                @lista.text = ""
                c_ctrl.input_data = []
            }
        }

        para ""

        flow {
            para " Wybierz przedział czasowy (YYYYCW): "
            flow {
                para " od: "
                @edit_od = edit_line "#{dh.get_current_yyyycw()}", :width => 120
                para " do: "
                @edit_do = edit_line "#{dh.get_current_yyyycw()}", :width => 120
            }
            



           
        }
        

        @moj_text = para @temp_txt
        
        


    }

    stack {
        flow {
            para " ASN? (MGO wymagane)"
            @checkbox_dla_asn = check
            
        }
        flow {
            para " "
            @guzik_dla_zrob_rqms = button "Generuj dla NB (Requirements)"
            @guzik_dla_zrob_sched = button "Generuj dla LA (Schedules)"
        }
        
        @log_text = para ""

    }
    

    @guzik_dla_zrob_rqms.click {
        
        my_input_data = []
        @moj_text.replace ""


        if @lista.text.size == 0 then
            @log_text.replace "LOG: text size: #{@lista.text.size} ? text: #{@lista.text} "
            alert "brak danych w liście!"
        else

            tmp_arr = @lista.text.split(/\n/)
            @log_text.replace tmp_arr[1]

            tmp_arr.each { |line|
                my_input_data << [ line.split(";")[0] , line.split(";")[1] , line.split(";")[2] ]
            }

            # @log_text.replace "LOG: tmp arr size: #{tmp_arr.size} ? text: #{c_ctrl.input_data}"
        end



        if my_input_data.size == 0 then
            alert "brak danych!"
        else


            t1 = @edit_od.text.to_f
            t2 = @edit_do.text.to_f
            
            #@log_text.replace "LOG: t1: #{t1} ? t2: #{t2} "

            if t1 < t2 || t1 == t2 then
                
                my_cws = [ t1.to_i.to_s, t2.to_i.to_s ]

                

                my_input_data.each { |element| 
                    element.each { |e| 
                        if e.nil? then e = "" end 
                    }
                }
                @moj_text.replace "Generating Requirements for: \n #{my_input_data} \n #{my_cws}"

                i_arr = my_input_data
                cw_arr = my_cws
                #c_ctrl.run_logic "SCHED", i_arr, cw_arr, false, false

                #$stderr.puts "RQMS: "
                #$stderr.puts my_cws


                

                #$stderr.puts @checkbox_dla_asn.checked? # OK # std bool
                is_asn = @checkbox_dla_asn.checked?
                run_rqm i_arr, cw_arr, "RQM", false, is_asn






                #$stderr.puts "after parsing..."
                @log_text.replace "PLEASE WAIT..."
                #$stderr.puts "finishing..."
                @log_text.replace "READY!"
                #$stderr.puts "READY!"
            else
                alert "niechronologicznie!"
            end
        end
    }

    @guzik_dla_zrob_sched.click {

        my_input_data = []
        @moj_text.replace ""



        if @lista.text.size == 0 then
            @log_text.replace "LOG: text size: #{@lista.text.size} ? text: #{@lista.text} "
            alert "brak danych w liście!"
        else

            tmp_arr = @lista.text.split(/\n/)
            @log_text.replace tmp_arr[1]

            tmp_arr.each { |line|

                # last duns sometimes catched as nil
                tmp = line.split(";")[2]
                if tmp.nil? then tmp = "" end

                my_input_data << [ line.split(";")[0] , line.split(";")[1] , tmp ]
            }

            # @log_text.replace "LOG: tmp arr size: #{tmp_arr.size} ? text: #{c_ctrl.input_data}"
        end


        if my_input_data.size == 0 then
            @log_text.replace "LOG: text size: #{@lista.text.size} ? text: #{@lista.text} "
            alert "brak danych w liście!"
        else

            t1 = @edit_od.text.to_f
            t2 = @edit_do.text.to_f

            #@log_text.replace "LOG: t1: #{t1} ? t2: #{t2} "

            if t1 < t2 || t1 == t2 then

                my_cws = [ t1.to_i.to_s, t2.to_i.to_s ]

                

                my_input_data.each { |element| 
                    element.each { |e| 
                        if e.nil? then e = "" end 
                    }
                }
                @moj_text.replace "Generating Sschedules for:  \n #{my_input_data} \n #{my_cws}"

                i_arr = my_input_data
                cw_arr = my_cws
                #c_ctrl.run_logic "SCHED", i_arr, cw_arr, false, false

                #$stderr.puts "SCHED: "
                #$stderr.puts my_cws


                

                #$stderr.puts @checkbox_dla_asn.checked? # OK std bool
                is_asn = @checkbox_dla_asn.checked?
                run_sched i_arr, cw_arr, "SCHED", false, is_asn






                #$stderr.puts "after parsing..."
                @log_text.replace "PLEASE WAIT..."
                #$stderr.puts wynik
                #$stderr.puts "finishing..."
                @log_text.replace "READY!"

               
            else
                alert "niechronologicznie!"
            end
        end
    }

    # button("Zdejmij Buciki!") { close() }


    @add.click {

        @lista.text += "" + @edit_plt.text + ";" + @edit_pn.text + ";" + @edit_duns.text + "\n"
        # c_ctrl.input_data << [ @edit_plt.text , @edit_pn.text , @edit_duns.text ]
    }
}
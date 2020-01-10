
class Time


    def self.inner_cwxx_calc cw
        if Time.new(Time.now.year,1,1).strftime('%W').to_i == 0 then
            cw += 1
        end

        to_return = ""
        if cw < 10 then
            to_return = "0#{cw}"
        else
            to_return = "#{cw}"
        end


        to_return
    end

    def self.cwxx
        cw = Time.now.strftime('%W').to_i
        r = Time.inner_cwxx_calc( cw )
        r
    end

    def cwxx
        cw = self.strftime('%W').to_i
        r = Time.inner_cwxx_calc( cw )
        r
    end


    def self.dzis
        rdzis = Time.new(Time.now.year, Time.now.month, Time.now.day, 11, 0, 0)
        rdzis
    end

    def dzis
        rdzis = Time.new(Time.now.year, Time.now.month, Time.now.day, 11, 0 , 0)
        rdzis
    end

    def self.jeden_dzien
        jd = 24 * 3600
        jd
    end
end

class DateHandler

    attr_accessor :ycw1, :ycw2

    def initialize
        @ycw1 = get_current_yyyycw()
        @ycw2 = get_current_yyyycw()
    end

    def get_current_yyyycw
        cw = Time.dzis.cwxx
        yyyy = Time.dzis.year
        yyyycw = "#{yyyy}#{cw}"
        yyyycw
    end



    def get_monday_from_ yyyycw

        yyyy = yyyycw[0...4]
        cw = yyyycw[4...6]


        dzis = Time.dzis
        dzis_weekday = dzis.wday
        moj_poniedzialek = dzis - (dzis_weekday * Time.jeden_dzien) + Time.jeden_dzien
        
        dzis_cw = Time.dzis.cwxx

        moj_poniedzialek_cw = moj_poniedzialek.cwxx        

        ten_poniedzialek = moj_poniedzialek
        ten_yyyycw = "#{moj_poniedzialek.year}00".to_i
        ten_yyyycw += moj_poniedzialek_cw.to_i




        if yyyy.to_i == moj_poniedzialek.year then

            diff_between_cw = cw.to_i - moj_poniedzialek_cw.to_i
            ten_poniedzialek = moj_poniedzialek + (diff_between_cw * 7 * Time.jeden_dzien)
        else
            # poniewaz odleglosci sa wieksze, niz na przestrzeni tylko jednego roku
            # nalezy nieco rozszerzyc 

            if yyyycw.to_i > ten_yyyycw
                

                until ten_yyyycw == yyyycw.to_i


                    # puts "" + ten_yyyycw.to_s + " " + ten_poniedzialek.to_s
                    
                    ten_poniedzialek += (7 * Time.jeden_dzien)
                    tpy = ten_poniedzialek.year
                    ten_yyyycw = "#{tpy}00".to_i
                    ten_yyyycw += ten_poniedzialek.cwxx.to_i
                end



            elsif yyyycw.to_i < ten_yyyycw


                until ten_yyyycw == yyyycw.to_i

                    # puts "" + ten_yyyycw.to_s + " " + ten_poniedzialek.to_s
                    
                    ten_poniedzialek -= (7 * Time.jeden_dzien)

                    # dodatkowa zmiana gdyby okazalo sie ze mamy znow 2018-12-31 jako cw1 dla roku 2019
                    tpy = ten_poniedzialek.year
                    ten_yyyycw = "#{tpy}00".to_i
                    ten_yyyycw += ten_poniedzialek.cwxx.to_i
                end

            end
        end

        ten_poniedzialek

    end


    def get_current_hour
        hh = Time.now.hour
        hh
    end
end





def test
    dh = DateHandler.new
    # puts dh.ycw1 + " " + dh.ycw2
    puts "xxx: #{dh.get_monday_from_("201850")}"
end

# test


def test2
    dh = DateHandler.new
    puts Time.cwxx
    puts Time.new(2019,7,1).cwxx
    puts Time.new(2019,1,1).cwxx
    puts dh.get_current_yyyycw
    puts dh.get_monday_from_ "202002" # 5 styczen 2020 powinien wyjsc
    puts dh.get_monday_from_ "201918"
end

#test2
require "google_drive"

# Creates a session. This will prompt the credential via command line for the
# first time and save it to config.json file for later usages.
# See this document to learn how to create config.json:
# https://github.com/gimite/google-drive-ruby/blob/master/doc/authorization.md



class Tabela
include Enumerable
    def initialize(kredencijali)
      @session = GoogleDrive::Session.from_config(kredencijali)
      @worksheet = @session.spreadsheet_by_key("1dqpqADKg94u9mENlkjEoTkd30zTvt-JbdYKG9wy7YQA").worksheets[0]
      napraviFunkcijeZaKolone()
    end
  
    def dvodimenzionalnaTabela()
      matrica = []
      (1..@worksheet.num_rows).each do |row|
        next if red_sadrzi_sumu?(row, ["total", "subtotal"])
        red = (1..@worksheet.num_cols).map do |col|
            kolona = @worksheet[row, col]
            kolona unless kolona.nil? || kolona.empty?
        end.compact
        matrica << red unless red.empty?
      end
      matrica
    end

    def red_sadrzi_sumu?(row, sume)
        (1..@worksheet.num_cols).any? do |col|
            element = @worksheet[row, col].to_s.downcase
            sume.any? { |suma| element.include?(suma) }
        end
    end

    def nadjiStartnuPoziciju()
        pocetniRed, pocetnaKolona = nil
    
        (1..@worksheet.num_rows).each do |row|
          (1..@worksheet.num_cols).each do |col|
            unless @worksheet[row, col].to_s.strip.empty?
              pocetniRed = row if pocetniRed.nil?
              pocetnaKolona = col if pocetnaKolona.nil? || col < pocetnaKolona
            end
          end
          break if pocetniRed && pocetnaKolona
        end
    
        [pocetniRed, pocetnaKolona]
    end

    def row(index)
        red = []
        
        pocetakRow, pocetakCol = nadjiStartnuPoziciju()
        (pocetakCol..@worksheet.num_cols).each do |col|
            red << @worksheet[pocetakRow+index, col] unless @worksheet[pocetakRow+index, col].empty?
        end

        red
    end

    def each(&block)
        pocetakRow, pocetakCol = nadjiStartnuPoziciju()
        (pocetakRow..@worksheet.num_rows).each do |row|
          (pocetakCol..@worksheet.num_cols).each do |col|
            cell_value = @worksheet[row, col]
            yield(cell_value) unless cell_value.empty? || cell_value.nil?
          end
        end
    end
    
    def indexKolone(imeKolone)
        pocetakRow, pocetakCol = nadjiStartnuPoziciju()
        (pocetakCol..@worksheet.num_cols).each do |col|
            return col if @worksheet[pocetakRow, col].to_s.strip.downcase==imeKolone.downcase
        end
        nil
    end

    def [](imeKolone)
        
        index = indexKolone(imeKolone)
        return nil if index.nil?
        kolona = []
        (1..@worksheet.num_rows).each do |row|
            element = @worksheet[row, index]
            kolona << element unless @worksheet[row, index].empty?
        end
        kolona
   
    end

    def napraviFunkcijeZaKolone()
        pocetakRow, pocetakCol = nadjiStartnuPoziciju()
        headers = (pocetakCol..@worksheet.num_cols).map { |col| @worksheet[pocetakRow, col].to_s[0].downcase + @worksheet[pocetakRow, col][1..-1].to_s.gsub(' ', '')}
        headers.each do |header|
          define_singleton_method(header) do
            kolona = []
            (pocetakRow+1..@worksheet.num_rows).each do |row|
              next if red_sadrzi_sumu?(row, ["total", "subtotal"])
              element = @worksheet[row, headers.find_index(header) + pocetakCol]
              kolona << element unless @worksheet[row, headers.find_index(header) + pocetakCol].empty?
            end
            def kolona.sum
                reduce(0){|sum,element| sum+element.to_f}
            end
      
            def kolona.avg
                sum / size.to_f
            end
            kolona
          end
        end
    end
    
  end

 

  


  def main
    t = Tabela.new("config.json")
  
    p t.dvodimenzionalnaTabela()

    p t.row(1)
    p t.row(1)[2]

    t.each do |element|
      p element
    end

   p t["Prva kolona"]
   p t["Prva kolona"][1]
  # izmena ne radi

   p t.prvaKolona
   p t.drugaKolona
   p t.trecaKolona
   p t.prvaKolona.sum
   p t.drugaKolona.avg
  # uzimanje pojedinacnog reda ne radi
  end
  
  main()
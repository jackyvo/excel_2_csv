require 'roo'
require 'csv'

class Excel2CSV

  def initialize file_path
    @file_path = file_path
  end

  def run
    parse
    export
  end

  private ##

  def parse
    xlsx = Roo::Spreadsheet.open(@file_path)

    # Verkaufsdaten
    verkaufsdaten_sheet = xlsx.sheet('Verkaufsdaten')
    verkaufsdaten_sheet_mapping = {}

    {
      '%Datum' => :date,
      'Transaktionszeit.Transactions' => :time,
      '%ArtNr' => :item_id,
      'Artikel.Artikel' => [:item_description, :name],
      'Warengruppencluster_neu.Artikel' => [:fine_category, :'product_category.name'],
      'Kostenstelle.Shops' => :store_map_id,
      'Umsatz' => :sold_price,
      'Menge' => :sold_quantity,
      'Retoure' => :waste_quantity
    }.map do |key, values|
      [*values].each { |item_key| verkaufsdaten_sheet_mapping[item_key] = key }
    end

    sheet = verkaufsdaten_sheet.sheet('Verkaufsdaten')
    
    @verkaufsdaten_rows = sheet.parse(verkaufsdaten_sheet_mapping)
    
    # Lieferdaten
    # @lieferdaten = xlsx.sheet('Lieferdaten')
  end

  def export
    # tranactions.csv
    verkaufsdaten_header = %i(date time item_id item_description fine_category store_map_id sold_price sold_quantity)
    write_to_csv('transactions.csv', verkaufsdaten_header, @verkaufsdaten_rows)

    # delivered.csv
  end

  def write_to_csv(file, header, rows)
    CSV.open("data/output/#{file}", "w") do |csv|
      csv << header

      rows.each do |row|
        csv << header.map { |key| formart_value(row[key], key) }
      end
    end
  end

  def formart_value value, key
    return nil if value == '-'

    formated = case key
    when :time
      value.strftime("%k:%M:%S").strip if value.is_a?(Date)
    when :date
      value.strftime("%Y-%m-%d") if value.is_a?(Date)
    when :sold_price
      '%.2f' % value
    when :sold_quantity
      '%.1f' % value
    when :item_description
      value.gsub(/^\d+\s/, '')
    end

    formated || value
  end
end
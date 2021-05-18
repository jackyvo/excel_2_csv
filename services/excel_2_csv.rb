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

    verkaufsdaten_sheet_mapping = {
      '%Datum' => :date,
      'Transaktionszeit.Transactions' => :time,
      '%ArtNr' => :item_id,
      'Artikel.Artikel' => [:item_description, :name],
      'Warengruppencluster_neu.Artikel' => [:fine_category, :'product_category.name'],
      'Kostenstelle.Shops' => :store_map_id,
      'Umsatz' => :sold_price,
      'Menge' => :sold_quantity,
      'Retoure' => :waste_quantity
    }.invert

    format_mapping!(verkaufsdaten_sheet_mapping)

    lieferdaten_sheet_mapping = {
      "Art.-Nr." => :item_id,
      "geliefert PE" => :delivered_quantity,
      "BäckereiNr" => :store_map_id,
      "Ebene 1\nLieferdatum" => :date,
      "Bezeichnung" => :item_description,
      "Mindestbestellmenge Lieferartikel" => :minimum_order
    }.invert

    format_mapping!(lieferdaten_sheet_mapping)


    # Verkaufsdaten
    sheet = xlsx.sheet('Verkaufsdaten')
    @verkaufsdaten_rows = sheet.parse(verkaufsdaten_sheet_mapping)
    
    # Lieferdaten
    sheet = xlsx.sheet('Lieferdaten')
    @lieferdaten_rows = sheet.parse(lieferdaten_sheet_mapping)
  end

  def export
    # tranactions.csv
    tranactions_header = %i(date time item_id item_description fine_category store_map_id sold_price sold_quantity)
    write_to_csv('transactions.csv', tranactions_header, @verkaufsdaten_rows)

    # delivered.csv
    delivered_header = %i(item_id item_description delivered_quantity store_map_id date)
    write_to_csv('delivered.csv', delivered_header, @lieferdaten_rows, "%y-%m-%d")
  end

  def write_to_csv(file, header, rows, date_format=nil)
    CSV.open("data/output/#{file}", "w") do |csv|
      csv << header

      rows.each do |row|
        csv << header.map { |key| formart_value(row[key], key, date_format) }
      end
    end
  end

  def formart_value value, key, format=nil
    return nil if value == '-' || value.nil?

    formated = case key
    when :time
      value.strftime("%k:%M:%S").strip if value.is_a?(Date)
    when :date
      value.strftime(format || "%Y-%m-%d") if value.is_a?(Date)
    when :sold_price, :delivered_quantity
      '%.2f' % value
    when :sold_quantity
      '%.1f' % value
    when :item_description
      value.gsub(/^\d+\s/, '')
    end

    formated || value
  end

  def format_mapping! mapping
    arrray_keys = mapping.keys.select { |key| key.is_a?(Array) }
    arrray_keys.each do |keys|
      keys.each do |key|
        mapping[key] = mapping[keys]
      end

      mapping.delete(keys)
    end
  end
end
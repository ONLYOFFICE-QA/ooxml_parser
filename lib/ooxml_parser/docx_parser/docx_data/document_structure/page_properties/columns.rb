require_relative 'columns/column'
module OoxmlParser
  class Columns
    attr_accessor :columns, :separator
    attr_accessor :count

    def initialize(columns_count)
      @count = columns_count
      @columns = []
    end

    # Parse Columns data
    # @param [Nokogiri::XML:Element] node with Columns data
    # @return [DocumentGrid] value of Columns data
    def self.parse(columns_grid)
      columns_count = 1
      columns_count = columns_grid.attribute('num').value.to_i unless columns_grid.attribute('num').nil?
      columns = Columns.new(columns_count)
      columns.separator = columns_grid.attribute('sep').value unless columns_grid.attribute('sep').nil?
      columns_grid.xpath('w:col').each do |col|
        width = (col.attribute('w').value.to_f / OoxmlParser.configuration.units_delimiter).round(2)
        space = (col.attribute('space').value.to_f / OoxmlParser.configuration.units_delimiter).round(2) unless col.attribute('space').nil?
        columns.columns << Column.new(width, space)
      end
      columns
    end
  end
end

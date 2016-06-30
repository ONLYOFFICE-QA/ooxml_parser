require_relative 'columns/column'
module OoxmlParser
  class Columns < OOXMLDocumentObject
    attr_accessor :separator
    attr_accessor :count
    attr_accessor :equal_width
    alias equal_width? equal_width
    attr_accessor :column_array

    def initialize(columns_count)
      @count = columns_count
      @column_array = []
    end

    # @return [Array, Column] accessor
    def [](key)
      @column_array[key]
    end

    # Parse Columns data
    # @param [Nokogiri::XML:Element] node with Columns data
    # @return [DocumentGrid] value of Columns data
    def self.parse(columns_grid)
      columns_count = 1
      columns_count = columns_grid.attribute('num').value.to_i unless columns_grid.attribute('num').nil?
      columns = Columns.new(columns_count)
      columns.separator = columns_grid.attribute('sep').value unless columns_grid.attribute('sep').nil?
      columns.equal_width = option_enabled?(columns_grid, 'equalWidth') unless columns_grid.attribute('equalWidth').nil?
      columns_grid.xpath('w:col').each do |col|
        columns.column_array << Column.parse(col)
      end
      columns
    end
  end
end

require_relative 'columns/column'
module OoxmlParser
  class Columns
    attr_accessor :columns, :separator
    attr_accessor :count

    def initialize(columns_count)
      @count = columns_count
      @columns = Array.new(@count)
    end

    # Parse Columns data
    # @param [Nokogiri::XML:Element] node with Columns data
    # @return [DocumentGrid] value of Columns data
    def self.parse(columns_grid)
      columns_count = 1
      columns_count = columns_grid.attribute('num').value.to_i unless columns_grid.attribute('num').nil?
      columns = Columns.new(columns_count)
      columns.separator = columns_grid.attribute('sep').value unless columns_grid.attribute('sep').nil?
      i = 0
      columns_grid.xpath('w:col').each do |col|
        width = col.attribute('w').value
        width = StringHelper.round(width.to_f / 566.9, 2) unless width.nil?
        space = col.attribute('space').value
        space = StringHelper.round(space.to_f / 566.9, 2) unless space.nil?
        columns.columns[i] = Column.new(width, space)
        i += 1
      end
      columns
    end
  end
end

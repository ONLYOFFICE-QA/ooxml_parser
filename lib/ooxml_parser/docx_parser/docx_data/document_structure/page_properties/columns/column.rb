module OoxmlParser
  # Class for parsing `w:col` object
  class Column
    attr_accessor :width, :space, :separator

    def initialize(width = nil, space = nil)
      @width = width
      @space = space
    end

    # Parse Column
    # @param [Nokogiri::XML:Node] node with Column
    # @return [Column] result of parsing
    def self.parse(node)
      column = Column.new
      node.attributes.each do |key, value|
        case key
        when 'w'
          column.width = (value.value.to_f / OoxmlParser.configuration.units_delimiter).round(2)
        when 'space'
          column.space = (value.value.to_f / OoxmlParser.configuration.units_delimiter).round(2)
        end
      end
      column
    end
  end
end

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
          column.width = OoxmlSize.new(value.value.to_f)
        when 'space'
          column.space = OoxmlSize.new(value.value.to_f)
        end
      end
      column
    end
  end
end

module OoxmlParser
  # Class for describing Shape Guide
  class ShapeGuide
    # @return [String] name of guide
    attr_accessor :name
    # @return [String] shape guide formula
    attr_accessor :formula

    # Parse Relationships
    # @param [Nokogiri::XML:Node] node with Shape Guide
    # @return [ShapeGuide] result of parsing
    def self.parse(node)
      guide = ShapeGuide.new
      node.attributes.each do |key, value|
        case key
        when 'name'
          guide.name = value.value.to_sym
        when 'fmla'
          guide.formula = value.value
        end
      end
      guide
    end
  end
end

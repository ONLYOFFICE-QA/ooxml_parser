module OoxmlParser
  class NumberingProperties < OOXMLDocumentObject
    attr_accessor :size, :font, :symbol, :start_at, :type, :ilvl, :numbering_properties

    def initialize(ilvl = 0, numbering_properties = nil)
      @ilvl = ilvl
      @numbering_properties = numbering_properties
    end

    def abstruct_numbering
      root_object.numbering.properties_by_num_id(@numbering_properties)
    end

    # Parse NumberingProperties
    # @param [Nokogiri::XML:Node] node with NumberingProperties
    # @param [OoxmlParser::OOXMLDocumentObject] parent parent object
    # @return [NumberingProperties] result of parsing
    def self.parse(node, parent)
      num_properties = NumberingProperties.new
      num_properties.parent = parent
      node.xpath('*').each do |num_prop_child|
        case num_prop_child.name
        when 'ilvl'
          num_properties.ilvl = num_prop_child.attribute('val').value.to_i
        when 'numId'
          num_properties.numbering_properties = num_prop_child.attribute('val').value.to_i
        end
      end
      num_properties
    end

    def numbering_level_current
      abstruct_numbering.level_list.each do |current_ilvl|
        return current_ilvl if current_ilvl.ilvl == @ilvl
      end
      nil
    end
  end
end

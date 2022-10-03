# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `numPr` tags
  class NumberingProperties < OOXMLDocumentObject
    attr_accessor :size, :font, :symbol, :start_at, :type, :image, :ilvl, :numbering_properties

    def initialize(ilvl = 0, parent: nil)
      @ilvl = ilvl
      super(parent: parent)
    end

    # @return [AbstractNumbering] AbstractNumbering of current properties
    def abstruct_numbering
      root_object.numbering.properties_by_num_id(@numbering_properties)
    end

    # Parse NumberingProperties
    # @param [Nokogiri::XML:Node] node with NumberingProperties
    # @return [NumberingProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'ilvl'
          @ilvl = node_child.attribute('val').value.to_i
        when 'numId'
          @numbering_properties = node_child.attribute('val').value.to_i
        end
      end
      self
    end

    # @return [AbstractNumbering] level list of current numbering
    def numbering_level_current
      abstruct_numbering.level_list.each do |current_ilvl|
        return current_ilvl if current_ilvl.ilvl == @ilvl
      end
      nil
    end
  end
end

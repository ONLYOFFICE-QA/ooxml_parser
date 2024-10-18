# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `numPr` tags
  class NumberingProperties < OOXMLDocumentObject
    attr_accessor :size, :font, :symbol, :start_at, :type, :image
    # @return [ValuedChild] i level
    attr_reader :i_level
    # @return [ValuedChild] numbering id
    attr_reader :num_id

    def initialize(ilvl = 0, parent: nil)
      @default_i_level = ilvl
      super(parent: parent)
    end

    # @return [AbstractNumbering] AbstractNumbering of current properties
    def abstruct_numbering
      root_object.numbering.properties_by_num_id(numbering_properties)
    end

    # Parse NumberingProperties
    # @param [Nokogiri::XML:Node] node with NumberingProperties
    # @return [NumberingProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'ilvl'
          @i_level = ValuedChild.new(:integer, parent: self).parse(node_child)
        when 'numId'
          @num_id = ValuedChild.new(:integer, parent: self).parse(node_child)
        end
      end
      self
    end

    # @return [Integer] numbering properties
    def numbering_properties
      @num_id.value
    end

    # @return [Integer] i-level value
    def ilvl
      return @default_i_level unless @i_level

      @i_level.value
    end

    # @param [Integer] i-level to find, current one by default
    # @return [AbstractNumbering] level list of current numbering
    def numbering_level_current(i_level = ilvl)
      abstruct_numbering.level_list.each do |current_ilvl|
        return current_ilvl if current_ilvl.ilvl == i_level
      end
      nil
    end
  end
end

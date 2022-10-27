# frozen_string_literal: true

module OoxmlParser
  # Class for `groupChr` data
  class GroupChar < OOXMLDocumentObject
    # @return [ValuedChild] symbol object
    attr_reader :symbol_object
    # @return [ValuedChild] position object
    attr_reader :position_object
    # @return [ValuedChild] vertical align object
    attr_reader :vertical_align_object
    attr_accessor :element

    # Parse GroupChar object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [GroupChar] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'groupChrPr'
          node_child.xpath('*').each do |node_child_child|
            case node_child_child.name
            when 'chr'
              @symbol_object = ValuedChild.new(:string, parent: self).parse(node_child_child)
            when 'pos'
              @position_object = ValuedChild.new(:string, parent: self).parse(node_child_child)
            when 'vertJc'
              @vertical_align_object = ValuedChild.new(:string, parent: self).parse(node_child_child)
            end
          end
        end
      end
      @element = DocxFormula.new(parent: self).parse(node)
      self
    end

    # @return [String] symbol value
    def symbol
      @symbol_object.value
    end

    # @return [String] position value
    def position
      @position_object.value
    end

    # @return [String] vertical align value
    def vertical_align
      @vertical_align_object.value
    end
  end
end

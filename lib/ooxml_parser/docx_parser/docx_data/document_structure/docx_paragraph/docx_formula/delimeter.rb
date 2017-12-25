module OoxmlParser
  # Class for `d` data
  class Delimiter < OOXMLDocumentObject
    attr_accessor :begin_character, :value, :end_character

    def initialize(parent: nil)
      @begin_character = '('
      @end_character = ')'
      @parent = parent
    end

    # Parse Delimiter object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Delimiter] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'dPr'
          node_child.xpath('*').each do |node_child_child|
            case node_child_child.name
            when 'begChr'
              @begin_character = ValuedChild.new(:string, parent: self).parse(node_child_child)
            when 'endChr'
              @end_character = ValuedChild.new(:string, parent: self).parse(node_child_child)
            end
          end
        end
      end
      @value = DocxFormula.new(parent: self).parse(node)
      self
    end
  end
end

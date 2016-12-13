module OoxmlParser
  # Class for `limUpp`, `limLow` data
  class Limit < OOXMLDocumentObject
    attr_accessor :type, :element, :limit

    def initialize(parent: nil)
      @type = :upper
      @parent = parent
    end

    # Parse Accent object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Accent] result of parsing
    def parse(node)
      @type = :lower if node.name == 'limLow'
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'lim'
          @limit = DocxFormula.new(parent: self).parse(node_child)
        end
      end
      @element = DocxFormula.new(parent: self).parse(node)
      self
    end
  end
end

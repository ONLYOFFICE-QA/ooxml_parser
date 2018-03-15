module OoxmlParser
  # Class for parsing `w:fldSimple` tags
  class FieldSimple < OOXMLDocumentObject
    # @return [String] instruction value
    attr_reader :instruction
    # @return [Array<ParagraphRun>] instruction value
    attr_reader :runs

    def initialize(parent: nil)
      @runs = []
      @parent = parent
    end

    # Parse FieldSimple object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [FieldSimple] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'instr'
          @instruction = value.value.to_s
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'r'
          @runs << ParagraphRun.new(parent: self).parse(node_child)
        end
      end
      self
    end

    # @return [True, False] is field simple page numbering
    def page_numbering?
      instruction.include?('PAGE')
    end
  end
end

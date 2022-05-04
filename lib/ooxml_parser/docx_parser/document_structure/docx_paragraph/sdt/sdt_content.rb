# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `w:w:sdtContent` tags
  class SDTContent < OOXMLDocumentObject
    # @return [Array <ParagraphRun, Table, ParagraphRun>] list of all elements in SDT
    attr_reader :elements

    def initialize(parent: nil)
      @elements = []
      super
    end

    # Parse SDTContent object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [SDTContent] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'p'
          @elements << DocxParagraph.new(parent: self).parse(node_child)
        when 'r'
          @elements << ParagraphRun.new(parent: self).parse(node_child)
        when 'tbl'
          @elements << Table.new(parent: self).parse(node_child)
        end
      end
      self
    end

    # @return [Array<DocxParagraphs>] list of paragraphs
    def paragraphs
      @elements.select { |obj| obj.is_a?(OoxmlParser::DocxParagraph) }
    end

    # @return [Array<ParagraphRun>] list of runs
    def runs
      @elements.select { |obj| obj.is_a?(OoxmlParser::ParagraphRun) }
    end

    # @return [Array<Table>] list of tables
    def tables
      @elements.select { |obj| obj.is_a?(OoxmlParser::Table) }
    end
  end
end

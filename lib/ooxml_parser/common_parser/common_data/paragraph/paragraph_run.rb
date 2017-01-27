require_relative 'paragraph_run/run_properties'
module OoxmlParser
  # Class for parsing `r` tags
  class ParagraphRun < OOXMLDocumentObject
    attr_accessor :properties, :text

    def initialize(properties = RunProperties.new, text = '', parent: nil)
      @properties = properties
      @text = text
      @parent = parent
    end

    # Parse ParagraphRun object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ParagraphRun] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'rPr'
          @properties = RunProperties.new(parent: self).parse(node_child)
        when 't'
          @text = node_child.text
        end
      end
      self
    end
  end
end

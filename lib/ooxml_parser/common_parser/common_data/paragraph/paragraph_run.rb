require_relative 'paragraph_run/run_properties'
module OoxmlParser
  # Class for parsing `r` tags
  class ParagraphRun < OOXMLDocumentObject
    attr_accessor :properties, :text

    def initialize(properties = RunProperties.new, text = '')
      @properties = properties
      @text = text
    end

    # Parse ParagraphRun object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [ParagraphRun] result of parsing
    def parse(character_node)
      character_node.xpath('*').each do |character_node_child|
        case character_node_child.name
        when 'rPr'
          @properties = RunProperties.new(parent: self).parse(character_node_child)
        when 't'
          @text = character_node_child.text
        end
      end
      self
    end
  end
end

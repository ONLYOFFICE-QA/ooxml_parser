# frozen_string_literal: true

require_relative 'paragraph/paragraph_properties'
require_relative 'paragraph/paragraph_run'
require_relative 'paragraph/text_field'
module OoxmlParser
  # Class for parsing `p` tags
  class Paragraph < OOXMLDocumentObject
    attr_accessor :properties, :runs, :text_field, :formulas
    # @return [AlternateContent] alternate content data
    attr_accessor :alternate_content

    def initialize(runs = [],
                   formulas = [],
                   parent: nil)
      @runs = runs
      @formulas = formulas
      @runs = []
      super(parent: parent)
    end

    alias characters runs
    alias character_style_array runs
    alias characters= runs=
    alias character_style_array= runs=

    # Parse Paragraph object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [Paragraph] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'pPr'
          @properties = ParagraphProperties.new(parent: self).parse(node_child)
        when 'fld'
          @text_field = TextField.new(parent: self).parse(node_child)
        when 'r'
          @runs << ParagraphRun.new(parent: self).parse(node_child)
        when 'AlternateContent'
          @alternate_content = AlternateContent.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end

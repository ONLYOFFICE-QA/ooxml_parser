# frozen_string_literal: true

require_relative 'xlsx_header_footer/header_footer_child'
module OoxmlParser
  # Class for parsing <headerFooter> tag
  class XlsxHeaderFooter < OOXMLDocumentObject
    # @return [Symbol] Specifies whether to align header with margins
    attr_reader :align_with_margins
    # @return [Symbol] Specifies whether first header is different
    attr_reader :different_first
    # @return [Symbol] Specifies whether odd and even headers are different
    attr_reader :different_odd_even
    # @return [Symbol] Specifies whether to scale header with document
    attr_reader :scale_with_document
    # @return [String] odd header
    attr_reader :odd_header
    # @return [String] odd footer
    attr_reader :odd_footer
    # @return [String] even header
    attr_reader :even_header
    # @return [String] even footer
    attr_reader :even_footer
    # @return [String] first header
    attr_reader :first_header
    # @return [String] first footer
    attr_reader :first_footer

    # Parse Header Footer data
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [XlsxHeaderFooter] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'alignWithMargins'
          @align_with_margins = attribute_enabled?(value)
        when 'differentFirst'
          @different_first = attribute_enabled?(value)
        when 'differentOddEven'
          @different_odd_even = attribute_enabled?(value)
        when 'scaleWithDoc'
          @scale_with_document = attribute_enabled?(value)
        end

        node.xpath('*').each do |node_child|
          case node_child.name
          when 'oddHeader'
            @odd_header = HeaderFooterChild.new(parent: parent, type: odd_header).parse(node_child)
          when 'oddFooter'
            @odd_footer = HeaderFooterChild.new(parent: parent, type: odd_footer).parse(node_child)
          when 'evenHeader'
            @even_header = HeaderFooterChild.new(parent: parent, type: even_header).parse(node_child)
          when 'evenFooter'
            @even_footer = HeaderFooterChild.new(parent: parent, type: even_footer).parse(node_child)
          when 'firstHeader'
            @first_header = HeaderFooterChild.new(parent: parent, type: first_header).parse(node_child)
          when 'firstFooter'
            @first_footer = HeaderFooterChild.new(parent: parent, type: first_footer).parse(node_child)
          end
        end
      end
      self
    end
  end
end

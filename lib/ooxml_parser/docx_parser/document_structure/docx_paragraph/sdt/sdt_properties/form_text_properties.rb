# frozen_string_literal: true

require_relative 'form_text_properties/form_text_comb'
require_relative 'form_text_properties/form_text_format'
module OoxmlParser
  # Class for parsing `textFormPr` tag
  class FormTextProperties < OOXMLDocumentObject
    # @return [True, False] specifies if field is multiline
    attr_reader :multiline
    # @return [True, False] specifies if size of field should be autofit
    attr_reader :autofit
    # @return [FormTextComb] parameters of text justification
    attr_reader :comb
    # @return [ValuedChild] characters limit
    attr_reader :maximum_characters
    # @return [FormTextFormat] text format
    attr_reader :format

    # Parse FormTextProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [FormTextProperties] result of parsing
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'multiLine'
          @multiline = option_enabled?(value)
        when 'autoFit'
          @autofit = option_enabled?(value)
        end
      end

      node.xpath('*').each do |node_child|
        case node_child.name
        when 'comb'
          @comb = FormTextComb.new(parent: self).parse(node_child)
        when 'maxCharacters'
          @maximum_characters = ValuedChild.new(:integer, parent: self).parse(node_child)
        when 'format'
          @format = FormTextFormat.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end

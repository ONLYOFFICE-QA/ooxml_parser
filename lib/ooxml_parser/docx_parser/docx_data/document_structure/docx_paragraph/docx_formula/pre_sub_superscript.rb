# frozen_string_literal: true

module OoxmlParser
  # Class for parsing `m:sPre` object
  class PreSubSuperscript < OOXMLDocumentObject
    # @return [DocxFormula] top value
    attr_accessor :top_value

    # @return [DocxFormula] bottom_value
    attr_accessor :bottom_value

    # @return [DocxFormula] base
    attr_accessor :base

    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'sub'
          @bottom_value = DocxFormula.new(parent: self).parse(node_child)
        when 'sup'
          @top_value = DocxFormula.new(parent: self).parse(node_child)
        when 'e'
          @base = DocxFormula.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end

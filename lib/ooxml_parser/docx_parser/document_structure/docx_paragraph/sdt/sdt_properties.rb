# frozen_string_literal: true

require_relative 'sdt_properties/form_properties'
require_relative 'sdt_properties/form_text_properties'
require_relative 'sdt_properties/combobox'
require_relative 'sdt_properties/dropdown_list'
require_relative 'sdt_properties/checkbox'
module OoxmlParser
  # Class for parsing `w:sdtPr` tags
  class SDTProperties < OOXMLDocumentObject
    # @return [ValuedChild] alias value
    attr_reader :alias
    # @return [ValuedChild] tag value
    attr_reader :tag
    # @return [ValuedChild] Locking Setting
    attr_reader :lock
    # @return [ComboBox] combobox
    attr_reader :combobox
    # @return [DropdownList] dropdown list
    attr_reader :dropdown_list
    # @return [CheckBox] checkbox
    attr_reader :checkbox
    # @return [FormProperties] form properties
    attr_reader :form_properties
    # @return [FormTextProperties] properties of text in form
    attr_reader :form_text_properties

    # Parse SDTProperties object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [SDTProperties] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'alias'
          @alias = ValuedChild.new(:string, parent: self).parse(node_child)
        when 'tag'
          @tag = ValuedChild.new(:string, parent: self).parse(node_child)
        when 'lock'
          @lock = ValuedChild.new(:symbol, parent: self).parse(node_child)
        when 'comboBox'
          @combobox = ComboBox.new(parent: self).parse(node_child)
        when 'dropDownList'
          @dropdown_list = DropdownList.new(parent: self).parse(node_child)
        when 'checkbox'
          @checkbox = CheckBox.new(parent: self).parse(node_child)
        when 'formPr'
          @form_properties = FormProperties.new(parent: self).parse(node_child)
        when 'textFormPr'
          @form_text_properties = FormTextProperties.new(parent: self).parse(node_child)
        end
      end
      self
    end
  end
end

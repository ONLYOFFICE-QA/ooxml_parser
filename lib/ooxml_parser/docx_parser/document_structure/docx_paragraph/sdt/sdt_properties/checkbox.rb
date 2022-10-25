# frozen_string_literal: true

require_relative 'checkbox/checkbox_state'
module OoxmlParser
  # Class for parsing `checkbox` tag
  class CheckBox < OOXMLDocumentObject
    # @return [True, False] specifies if checkbox is checked
    attr_reader :checked
    # @return [CheckBoxState] options of checked state
    attr_reader :checked_state
    # @return [CheckBoxState] options of unchecked state
    attr_reader :unchecked_state
    # @return [ValuedChild] group key
    attr_reader :group_key

    # Parse CheckBox object
    # @param node [Nokogiri::XML:Element] node to parse
    # @return [CheckBox] result of parsing
    def parse(node)
      node.xpath('*').each do |node_child|
        case node_child.name
        when 'checked'
          @checked = option_enabled?(node_child, 'hidden')
        when 'checkedState'
          @checked_state = CheckBoxState.new(parent: self).parse(node_child)
        when 'uncheckedState'
          @unchecked_state = CheckBoxState.new(parent: self).parse(node_child)
        when 'groupKey'
          @group_key = ValuedChild.new(:string, parent: self).parse(node_child)
        end
      end
      self
    end
  end
end

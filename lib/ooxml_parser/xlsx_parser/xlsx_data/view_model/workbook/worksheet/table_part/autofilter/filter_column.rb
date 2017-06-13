module OoxmlParser
  # Class for `filterColumn` data
  # The filterColumn collection identifies a particular
  # column in the AutoFilter range and specifies
  # filter information that has been applied to this column.
  class FilterColumn < OOXMLDocumentObject
    # @return [True, False] Flag indicating whether the filter button is visible.
    attr_accessor :show_button

    def initialize(parent: nil)
      @show_button = true
      @parent = parent
    end

    # Parse FilterColumn data
    # @param [Nokogiri::XML:Element] node with FilterColumn data
    # @return [FilterColumn] value of FilterColumn data
    def parse(node)
      node.attributes.each do |key, value|
        case key
        when 'showButton'
          @show_button = attribute_enabled?(value)
        end
      end
      self
    end
  end
end

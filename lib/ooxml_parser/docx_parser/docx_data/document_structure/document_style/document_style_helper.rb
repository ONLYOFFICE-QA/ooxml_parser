module OoxmlParser
  # Helper methods for working with Style List
  module DocumentStyleHelper
    # Return document style by its name
    # @param name [String] name of style
    # @return [DocumentStyle, nil]
    def document_style_by_name(name)
      @document_styles.each do |style|
        return style if style.name == name
      end
      nil
    end

    # Return document style by its id
    # @param name [String] id of style
    # @return [DocumentStyle, nil]
    def document_style_by_id(id)
      @document_styles.each do |style|
        return style if style.style_id == id
      end
      nil
    end

    # Check if style exists in current document
    # @param name [String] name of style
    # @return [True, False]
    def style_exist?(name)
      !document_style_by_name(name).nil?
    end
  end
end

# Style Parameter Data
module OoxmlParser
  class StyleParametres
    attr_accessor :q_format, :hidden, :name

    def initialize(name = nil, q_format = false, hidden = false)
      @name = name
      @q_format = q_format
      @hidden = hidden
    end

    def ==(other)
      all_instance_variables = instance_variables
      all_instance_variables.each do |current_attributes|
        unless instance_variable_get(current_attributes) == other.instance_variable_get(current_attributes)
          return false
        end
      end
      true
    end

    def self.parse(style_tag)
      style = StyleParametres.new
      style_tag.xpath('w:name').each do |name|
        style.name = name.attribute('val').value
      end
      style_tag.xpath('w:qFormat').each do |q_format|
        next if q_format.attribute('val').nil?
        next if q_format.attribute('val').value == false || q_format.attribute('val').value == 'off' || q_format.attribute('val').value == '0'
        style.q_format = true
      end
      style
    end
  end
end

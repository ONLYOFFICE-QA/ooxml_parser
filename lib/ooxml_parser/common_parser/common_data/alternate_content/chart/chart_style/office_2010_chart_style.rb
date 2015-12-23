# XLSX 2010 Chart Style
module OoxmlParser
  class Office2010ChartStyle
    attr_accessor :style_number

    def self.parse(chart_style_node)
      chart_style = Office2010ChartStyle.new
      chart_style_node.xpath('*').each do |chart_style_node_child|
        case chart_style_node_child.name
        when 'style'
          chart_style.style_number = chart_style_node_child.attribute('val').value.to_i
        end
      end
      chart_style
    end
  end
end

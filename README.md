# SheetBuilder

Sheet Builder is an abstraction of the axlsx gem that allows you to quickly make and spreadsheet templates that exists as
arrays of Ruby hashes for easy spreadsheet generation and data input.

TODO:

- Clean up styling options

## Installation

Add this line to your application's Gemfile:

```ruby
gem 'sheet_builder'
```

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install sheet_builder

## Usage

This gem revolves around the Blueprint class, which you supply with these (optional) component types:
- Elements: Components with specified row and column (i.e. headers, etc)
- Titles: Components that are laid out automagically in order, and used for placement/retrieval of data
  - Column titles span horizontally, Row titles span vertically
- Title Starting Points: Where your titles will start being placed in (defaults to first row and first column)

Start by creating a package, workbook, and sheet with Axlsx

```ruby
package = Axlsx::Package.new
workbook = package.workbook
sheet = workbook.add_worksheet name: 'My Awesome Sheet'
```

Then create a Blueprint object with your desired Components
```ruby
elements = [
  {text: 'This is a simple element!', row: 1, col: 1},
  {text: 'This is a complex element', row: 2, col: 2, style: [:bold, :center], color: 'FD7F80', borders: true, merge: 2}
]

column_titles = [
  {text: 'This is a simple column_title!'},
  {text: 'This is a colorful column_title!', color: '9BCBFD'},
  {text: 'This column_title has a comment!', comment: 'Told you so!'},
  {text: 'This column_title has a dropdown list for its options!', list: ['told', 'you', 'so']}
]
```

Then create a new blueprint object with these components. You can specific where
titles will start with column/row_titles_start: [row, col]

```ruby
blueprint = SheetBuilder::Blueprint.new elements: elements,
                                        column_titles: column_titles,
                                        column_titles_start: [3,1]
```

Then call build! on the blueprint, and pass it a sheet, and optional data

```ruby
column_data = [
  {text: 'I will always be placed under my correct title!', title: 'This is a simple column_title!'},
  {text: 'Me too!', title: 'This is a simple column_title!'},
  {text: 'Me three!', title: 'This column_title has a comment!'}
]

blueprint.build! sheet: sheet, column_data: column_data
```

Finally, you can save your workbook!

```ruby
package.serialize 'test.xlsx'
```

## Contributing

Bug reports and pull requests are welcome on GitHub at https://github.com/[USERNAME]/sheet_builder. This project is intended to be a safe, welcoming space for collaboration, and contributors are expected to adhere to the [Contributor Covenant](http://contributor-covenant.org) code of conduct.


## License

The gem is available as open source under the terms of the [MIT License](http://opensource.org/licenses/MIT).

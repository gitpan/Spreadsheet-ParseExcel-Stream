package Spreadsheet::ParseExcel::Stream;

use strict;
use warnings;

use Spreadsheet::ParseExcel;
use Scalar::Util qw(weaken);
use Coro;

our $VERSION = '0.02';

sub new {
  my ($class, $file) = @_;

  my $main = Coro::State->new();
  my ($xls,$parser);

  my ($wb, $idx, $row, $col, $cell);
  my $tmp = my $handler = sub {
    ($wb, $idx, $row, $col, $cell) = @_;
    $parser->transfer($main);
  };

  my $tmp_p = $parser = Coro::State->new(sub {
    $xls->Parse($file);
    # Flag the generator that we're done
    undef $xls;
    # If we don't transfer back when done parsing,
    # it's an implicit program exit (oops!)
    $parser->transfer($main)
  });
  weaken($parser);

  $xls = Spreadsheet::ParseExcel->new(
    CellHandler => $handler,
    NotSetCell => 1,
  );

  # Returns the next cell of the spreadsheet
  my $generator = sub {

    # Just in case we ask for the next cell when we're already done
    return unless $xls;

    $main->transfer($parser);
    return [ $wb, $idx, $row, $col, $cell ] if $xls;

    # We're done with these threads
    $main->cancel();
    $parser->cancel();
    return;
  };
  my $nxt_cell = $generator->();

  bless {
    # Save a reference to the parser so it doesn't disappear
    # until the object is destroyed.
    PARSER => $tmp_p,
    NEXT_CELL => $nxt_cell,
    SUB      => $generator,
  }, $class . '::Sheet';
}

package Spreadsheet::ParseExcel::Stream::Sheet;

sub sheet {
  my $self = shift;
  return unless $self->{NEXT_CELL};
  return $self;
}

sub worksheet {
  my $self = shift;
  my $row = $self->{NEXT_CELL};
  my $wb = $row->[0];
  return $wb->worksheet($row->[1]);
}

sub name {
  my $self = shift;
  return $self->worksheet()->{Name};
}

sub next_row {
  my ($self, $current) = @_;
  return $self->{CURR_ROW} if $current;

  return $self->{NEW_WS} = 0 if $self->{NEW_WS};

  # Save original cell so we can detect change in worksheet
  my $orig_cell = my $curr_cell = $self->{NEXT_CELL};
  my $f = $self->{SUB};

  # Initialize row with first cell
  my @row = ();
  $row[ $curr_cell->[3] ] = $curr_cell;
  my $nxt_cell = $f->();

  # Collect current row on current worksheet
  while ( $nxt_cell && $nxt_cell->[1] == $curr_cell->[1] && $nxt_cell->[2] == $curr_cell->[2] ) {
    $curr_cell = $nxt_cell;
    $row[ $curr_cell->[3] ] = $curr_cell;
    $nxt_cell = $f->();
  }
  $self->{NEXT_CELL} = $nxt_cell;
  $self->{NEW_WS}++ if !$nxt_cell || $orig_cell->[1] != $nxt_cell->[1];
  return $self->{CURR_ROW} = \@row;
}

sub row {
  my ($self, $current) = @_;
  unless ($current) {
    my $row = $self->next_row();
    return unless $row;
  }
  return [ map { defined $_ ? $_->[4]->value() : $_ } @{$self->{CURR_ROW}} ];
}

sub row_unformatted {
  my ($self, $current) = @_;
  unless ($current) {
    my $row = $self->next_row();
    return unless $row;
  }
  return [ map { defined $_ ? $_->[4]->unformatted() : $_ } @{$self->{CURR_ROW}} ];
}

1;

__END__

=head1 NAME

Spreadsheet::ParseExcel::Stream - Simple interface to Excel data with no memory overhead

=head1 SYNOPSIS

  my $xls = Spreadsheet::ParseExcel::Stream->new($xls_file);
  while ( my $sheet = $xls->sheet() ) {
    while ( my $row = $sheet->row ) {
      my @data = @$row;
    }
  }

=head1 DESCRIPTION

A simple iterative interface to L<Spreadsheet::ParseExcel>, similar to L<Spreadsheet::ParseExcel::Simple>,
but does not parse the entire document to memory. Uses the hints provided in the L<Spreadsheet::ParseExcel>
docs to reduce memory usage, and returns the data row by row and sheet by sheet.

=head1 METHODS

=head2 new

  my $xls = Spreadsheet::ParseExcel::Stream->new($xls_file);

Opens the spreadsheet and returns an object to iterate through the data.

=head2 sheet

Returns the sheet of the next cell of the spreadsheet.

=head2 row

Returns the next row of data from the current spreadsheet. The data is the formatted
contents of each cell as returned by the $cell->value() method of Spreadsheet::ParseExcel.

If a true argument is passed in, returns the current row of data without advancing to the
next row.

=head2 unformatted

Returns the next row of data from the current spreadsheet as returned
by the $cell->unformatted() method of Spreadsheet::ParseExcel.

If a true argument is passed in, returns the current row of data without advancing to the
next row.

=head2 next_row

Returns the next row of cells from the current spreadsheet as Spreadsheet::ParseExcel
cell objects.

If a true argument is passed in, returns the current row without advancing to the
next row.

=head2 name

Returns the name of the next cell of the spreadsheet.

=head2 worksheet

Returns the worksheet containing the next cell of data as a Spreadsheet::ParseExcel object.

=head1 AUTHOR

Douglas Wilson, E<lt>dougw@cpan.org<gt>

=head1 COPYRIGHT AND LICENSE

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.

=head1 SEE ALSO

L<Spreadsheet::ParseExcel>, L<Spreadsheet::ParseExcel::Simple>

=cut

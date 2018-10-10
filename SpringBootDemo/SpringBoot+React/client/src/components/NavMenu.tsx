import * as React from 'react';
import { Link, NavLink } from 'react-router-dom';

export class NavMenu extends React.Component<{}, {}> {
    public render() {
        return <div className='main-nav'>
                <div className='navbar navbar-inverse'>
                <div className='navbar-header'>
                    <button type='button' className='navbar-toggle' data-toggle='collapse' data-target='.navbar-collapse'>
                        <span className='sr-only'>Toggle navigation</span>
                        <span className='icon-bar'></span>
                        <span className='icon-bar'></span>
                        <span className='icon-bar'></span>
                    </button>
                    <Link className='navbar-brand' to={ '/' }>GrapeCity Documents for Excel, Java Edition</Link>
                </div>
                <div className='clearfix'></div>
                <div className='navbar-collapse collapse'>
                    <ul className='nav navbar-nav'>
                        <li>
                            <NavLink to={ '/' } exact activeClassName='active'>
                                <span className='glyphicon glyphicon-home'></span> Home
                            </NavLink>
                        </li>
                        <li>
                            <NavLink to={ '/ExcelTemplateDemo' } activeClassName='active'>
                                <span className='glyphicon glyphicon-th-list'></span> Excel Template Demo
                            </NavLink>
                        </li>
                        <li>
                            <NavLink to={ '/ProgrammingDemo' } activeClassName='active'>
                                <span className='glyphicon glyphicon-th-list'></span> Programming API Demo
                            </NavLink>
                        </li>
                        <li>
                            <NavLink to={'/ExcelIODemo'} activeClassName='active'>
                                <span className='glyphicon glyphicon-th-list'></span> Excel IO Demo
                            </NavLink>
                        </li>
                    </ul>
                </div>
            </div>
        </div>;
    }
}

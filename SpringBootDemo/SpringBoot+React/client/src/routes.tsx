import * as React from 'react';
import { Route } from 'react-router-dom';
import { Layout } from './components/Layout';
import { Home } from './components/Home';
import { ExcelTemplateDemo } from './components/ExcelTemplateDemo';
import { ProgrammingDemo } from './components/ProgrammingDemo';
import { ExcelIODemo } from './components/ExcelIODemo';

export const routes = <Layout>
    <Route exact path='/' component={ Home } />
    <Route path='/ExcelTemplateDemo' component={ExcelTemplateDemo} />
    <Route path='/ProgrammingDemo' component={ProgrammingDemo} />
    <Route path='/ExcelIODemo' component={ ExcelIODemo } />
</Layout>;
